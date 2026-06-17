use std::fmt::{Display, Formatter};
use std::io::{Read, Write};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WebSocketFrame {
    Text(String),
    Close,
    Ping(Vec<u8>),
    Pong,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebSocketProtocolError {
    pub close_code: u16,
    pub reason: String,
}

impl WebSocketProtocolError {
    fn protocol(reason: impl Into<String>) -> Self {
        Self {
            close_code: 1002,
            reason: reason.into(),
        }
    }

    fn too_large(reason: impl Into<String>) -> Self {
        Self {
            close_code: 1009,
            reason: reason.into(),
        }
    }
}

#[derive(Debug)]
pub enum WebSocketCodecError {
    Io(std::io::Error),
    Protocol(WebSocketProtocolError),
}

impl Display for WebSocketCodecError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Protocol(error) => formatter.write_str(&error.reason),
        }
    }
}

impl std::error::Error for WebSocketCodecError {}

impl From<std::io::Error> for WebSocketCodecError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

pub struct WebSocketCodec;

impl WebSocketCodec {
    /// Reads one client WebSocket frame.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed frames, unmasked client frames, oversized
    /// payloads, invalid UTF-8 text, or underlying I/O failures.
    pub fn read_frame(
        stream: &mut impl Read,
        max_frame_bytes: usize,
    ) -> Result<Option<WebSocketFrame>, WebSocketCodecError> {
        let mut header = [0_u8; 2];
        if let Err(error) = stream.read_exact(&mut header) {
            return if error.kind() == std::io::ErrorKind::UnexpectedEof {
                Ok(None)
            } else {
                Err(WebSocketCodecError::Io(error))
            };
        }
        let fin = header[0] & 0x80 != 0;
        let opcode = header[0] & 0x0f;
        let masked = header[1] & 0x80 != 0;
        if !fin {
            return Err(WebSocketCodecError::Protocol(
                WebSocketProtocolError::protocol("Fragmented WebSocket frames are not supported."),
            ));
        }
        let mut length = usize::from(header[1] & 0x7f);
        if length == 126 {
            let mut extended = [0_u8; 2];
            stream.read_exact(&mut extended)?;
            length = usize::from(u16::from_be_bytes(extended));
        } else if length == 127 {
            let mut extended = [0_u8; 8];
            stream.read_exact(&mut extended)?;
            let raw = u64::from_be_bytes(extended);
            length = raw.try_into().map_err(|_| {
                WebSocketCodecError::Protocol(WebSocketProtocolError::too_large(
                    "WebSocket frame is too large.",
                ))
            })?;
        }
        if length > max_frame_bytes {
            return Err(WebSocketCodecError::Protocol(
                WebSocketProtocolError::too_large("WebSocket frame exceeds configured byte limit."),
            ));
        }
        let mut mask = [0_u8; 4];
        if masked {
            stream.read_exact(&mut mask)?;
        } else if matches!(opcode, 0x1 | 0x8 | 0x9 | 0xA) {
            return Err(WebSocketCodecError::Protocol(
                WebSocketProtocolError::protocol("Client WebSocket frames must be masked."),
            ));
        }
        let mut payload = vec![0_u8; length];
        stream.read_exact(&mut payload)?;
        if masked {
            for (index, byte) in payload.iter_mut().enumerate() {
                *byte ^= mask[index % 4];
            }
        }
        match opcode {
            0x1 => String::from_utf8(payload)
                .map(WebSocketFrame::Text)
                .map(Some)
                .map_err(|_| {
                    WebSocketCodecError::Protocol(WebSocketProtocolError::protocol(
                        "Invalid UTF-8 text frame.",
                    ))
                }),
            0x8 => Ok(Some(WebSocketFrame::Close)),
            0x9 => Ok(Some(WebSocketFrame::Ping(payload))),
            0xA => Ok(Some(WebSocketFrame::Pong)),
            _ => Err(WebSocketCodecError::Protocol(
                WebSocketProtocolError::protocol(format!("Unsupported WebSocket opcode {opcode}.")),
            )),
        }
    }

    /// Writes one unmasked server text frame.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream cannot be written or flushed.
    pub fn write_text(stream: &mut impl Write, text: &str) -> Result<(), WebSocketCodecError> {
        Self::write_frame(stream, 0x1, text.as_bytes())
    }

    /// Writes one unmasked server pong frame.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream cannot be written or flushed.
    pub fn write_pong(stream: &mut impl Write, payload: &[u8]) -> Result<(), WebSocketCodecError> {
        Self::write_frame(stream, 0xA, payload)
    }

    /// Writes one unmasked server close frame.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream cannot be written or flushed.
    pub fn write_close(
        stream: &mut impl Write,
        code: u16,
        reason: &str,
    ) -> Result<(), WebSocketCodecError> {
        let reason = reason.as_bytes();
        let reason = &reason[..reason.len().min(123)];
        let mut payload = Vec::with_capacity(2 + reason.len());
        payload.extend_from_slice(&code.to_be_bytes());
        payload.extend_from_slice(reason);
        Self::write_frame(stream, 0x8, &payload)
    }

    fn write_frame(
        stream: &mut impl Write,
        opcode: u8,
        payload: &[u8],
    ) -> Result<(), WebSocketCodecError> {
        let mut header = vec![0x80 | opcode];
        if payload.len() < 126 {
            header.push(payload.len().try_into().unwrap_or(125));
        } else if let Ok(length) = u16::try_from(payload.len()) {
            header.push(126);
            header.extend_from_slice(&length.to_be_bytes());
        } else {
            header.push(127);
            header.extend_from_slice(&(payload.len() as u64).to_be_bytes());
        }
        stream.write_all(&header)?;
        stream.write_all(payload)?;
        stream.flush()?;
        Ok(())
    }
}

#[cfg(test)]
#[path = "websocket_codec_tests.rs"]
mod websocket_codec_tests;
