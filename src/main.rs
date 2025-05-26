use std::{
    io::{self, Read as _},
    thread,
};

use util::TerminationFlag;
use windows::{
    core::Result,
    Win32::System::Com::{CoInitialize, CoUninitialize},
};

mod record;
mod util;

fn main() -> Result<()> {
    let notifier = TerminationFlag::default();
    let runner = thread::spawn({
        let notifier = notifier.clone();
        || unsafe {
            CoInitialize(None).ok()?;
            let result = record::to_stdout(notifier);
            CoUninitialize();
            result
        }
    });
    io::stdin().read(&mut []).unwrap();
    notifier.notify();
    runner.join().unwrap()
}
