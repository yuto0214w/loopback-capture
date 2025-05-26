use std::io;

use windows::{
    core::Result,
    Win32::{
        Media::{
            Audio::{
                eConsole, eRender, IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator,
                MMDeviceEnumerator, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK,
                WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVE_FORMAT_PCM,
            },
            KernelStreaming::{KSDATAFORMAT_SUBTYPE_PCM, WAVE_FORMAT_EXTENSIBLE},
            Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT},
        },
        System::{
            Com::{CoCreateInstance, CoTaskMemFree, CLSCTX_ALL},
            Threading::Sleep,
        },
    },
};

use crate::util::TerminationFlag;

pub fn to_stdout(notifier: TerminationFlag) -> Result<()> {
    _record_internal(&mut io::stdout().lock(), notifier)
}

// References:
// https://learn.microsoft.com/en-us/windows/win32/coreaudio/capturing-a-stream
// https://github.com/mvaneerde/blog/blob/28f8cbdfbcbc8b61a13a710e418b55c2f8c7600c/loopback-capture/loopback-capture/loopback-capture.cpp

const REFTIMES_PER_MILLISEC: u32 = 10000;
const REFTIMES_PER_SEC: u32 = REFTIMES_PER_MILLISEC * 1000;

fn _record_internal<W: io::Write>(writer: &mut W, notifier: TerminationFlag) -> Result<()> {
    let mut pwfx: Option<*mut WAVEFORMATEX> = None;

    let result: Result<()> = (|| unsafe {
        let requested_duration = REFTIMES_PER_SEC;

        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let device = enumerator.GetDefaultAudioEndpoint(eRender, eConsole)?;
        let audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
        let pwfx = *pwfx.get_or_insert(audio_client.GetMixFormat()?);

        // coerce to s16le
        match (*pwfx).wFormatTag as u32 {
            WAVE_FORMAT_IEEE_FLOAT => {
                (*pwfx).wFormatTag = WAVE_FORMAT_PCM as u16;
                (*pwfx).wBitsPerSample = 16;
                (*pwfx).nBlockAlign = (*pwfx).nChannels * (*pwfx).wBitsPerSample / 8;
                (*pwfx).nAvgBytesPerSec = (*pwfx).nBlockAlign as u32 * (*pwfx).nSamplesPerSec;
            }
            WAVE_FORMAT_EXTENSIBLE => {
                let pwfex: *mut WAVEFORMATEXTENSIBLE = std::mem::transmute(pwfx);
                if { (*pwfex).SubFormat } == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
                    (*pwfex).SubFormat = KSDATAFORMAT_SUBTYPE_PCM;
                    (*pwfex).Samples.wValidBitsPerSample = 16;
                    (*pwfx).wBitsPerSample = 16;
                    (*pwfx).nBlockAlign = (*pwfx).nChannels * (*pwfx).wBitsPerSample / 8;
                    (*pwfx).nAvgBytesPerSec = (*pwfx).nBlockAlign as u32 * (*pwfx).nSamplesPerSec;
                }
            }
            _ => panic!("not sure how to treat this format"),
        }

        audio_client.Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            AUDCLNT_STREAMFLAGS_LOOPBACK,
            requested_duration as i64,
            0,
            pwfx,
            None,
        )?;

        let buffer_frame_count = audio_client.GetBufferSize()?;
        let capture_client: IAudioCaptureClient = audio_client.GetService()?;

        let actual_duration = REFTIMES_PER_SEC * buffer_frame_count / pwfx.read().nSamplesPerSec;

        audio_client.Start()?;

        let mut data: *mut u8 = std::ptr::null_mut();
        let mut frames_read: u32 = 0;
        let mut flags: u32 = 0;
        while !notifier.should_terminate() {
            Sleep(actual_duration / REFTIMES_PER_MILLISEC / 2);

            let mut packet_length = capture_client.GetNextPacketSize()?;
            while packet_length != 0 {
                capture_client.GetBuffer(&mut data, &mut frames_read, &mut flags, None, None)?;

                let data_size = frames_read as usize * (*pwfx).nBlockAlign as usize;
                writer
                    .write_all(std::slice::from_raw_parts(data, data_size))
                    .unwrap();

                capture_client.ReleaseBuffer(frames_read)?;
                packet_length = capture_client.GetNextPacketSize()?;
            }
        }

        audio_client.Stop()
    })();

    unsafe {
        CoTaskMemFree(pwfx.map(|pwfx| pwfx as *const _));
    }
    result
}
