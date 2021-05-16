Com-rs update to version 0.3 or 0.4 is blocked by the fact that it breaks the
rust-analyzer. (Reason is that Rust analyzer is unable to support macros with
same name as the macro.)

Also the future of com-rs is in limbo ( https://github.com/microsoft/com-rs/issues/206 ) there is no point to use time to migrate with sake of migration at this moment.

## Other

com_interfaces -> uuid

pub trait ->
pub unsafe interface

ComRc<dyn (.+?)>
$1
