#[must_use]
pub(crate) fn same_major_version(left: &str, right: &str) -> bool {
    left.split('.').next() == right.split('.').next()
}

#[cfg(test)]
#[path = "addin_protocol_version_tests.rs"]
mod addin_protocol_version_tests;
