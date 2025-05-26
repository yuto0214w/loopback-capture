use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};

#[derive(Default)]
pub struct TerminationFlag(Arc<AtomicBool>);

impl Clone for TerminationFlag {
    fn clone(&self) -> Self {
        Self(Arc::clone(&self.0))
    }
}

impl TerminationFlag {
    pub fn notify(&self) {
        self.0.store(true, Ordering::Release);
    }

    pub fn should_terminate(&self) -> bool {
        self.0.load(Ordering::Acquire)
    }
}
