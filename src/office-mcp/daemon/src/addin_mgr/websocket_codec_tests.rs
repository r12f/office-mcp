use super::{WebSocketCodec, WebSocketCodecError, WebSocketFrame};

#[test]
fn reads_masked_text_frame_and_writes_text_frame() {
    let input = masked_text_frame("hello");
    let frame = WebSocketCodec::read_frame(&mut input.as_slice(), 1024).expect("frame");
    assert_eq!(frame, Some(WebSocketFrame::Text("hello".to_string())));

    let mut output = Vec::new();
    WebSocketCodec::write_text(&mut output, "ok").expect("write");
    assert_eq!(output, vec![0x81, 0x02, b'o', b'k']);
}

#[test]
fn maps_protocol_errors_to_close_codes() {
    let unmasked_text = vec![0x81, 0x02, b'h', b'i'];
    let error = WebSocketCodec::read_frame(&mut unmasked_text.as_slice(), 1024)
        .expect_err("unmasked client frame rejected");
    assert_ws_error(error, 1002, "masked");

    let fragmented_text = {
        let mut frame = masked_text_frame("hi");
        frame[0] = 0x01;
        frame
    };
    let error = WebSocketCodec::read_frame(&mut fragmented_text.as_slice(), 1024)
        .expect_err("fragmented frame rejected");
    assert_ws_error(error, 1002, "Fragmented");

    let binary_frame = {
        let mut frame = masked_text_frame("hi");
        frame[0] = 0x82;
        frame
    };
    let error = WebSocketCodec::read_frame(&mut binary_frame.as_slice(), 1024)
        .expect_err("unsupported opcode rejected");
    assert_ws_error(error, 1002, "Unsupported");

    let oversized = masked_text_frame("hello");
    let error = WebSocketCodec::read_frame(&mut oversized.as_slice(), 2)
        .expect_err("oversized frame rejected");
    assert_ws_error(error, 1009, "exceeds");
}

#[test]
fn writes_close_frame_with_code_and_reason() {
    let mut output = Vec::new();
    WebSocketCodec::write_close(&mut output, 4002, "Heartbeat timeout").expect("close");

    assert_eq!(output[0], 0x88);
    assert_eq!(output[1] as usize, 2 + "Heartbeat timeout".len());
    assert_eq!(u16::from_be_bytes([output[2], output[3]]), 4002);
    assert_eq!(&output[4..], b"Heartbeat timeout");
}

fn assert_ws_error(error: WebSocketCodecError, close_code: u16, reason: &str) {
    match error {
        WebSocketCodecError::Protocol(error) => {
            assert_eq!(error.close_code, close_code);
            assert!(
                error.reason.contains(reason),
                "expected `{}` to contain `{}`",
                error.reason,
                reason
            );
        }
        other => panic!("expected websocket protocol error, got {other:?}"),
    }
}

fn masked_text_frame(text: &str) -> Vec<u8> {
    let mask = [1_u8, 2, 3, 4];
    let mut frame = vec![0x81];
    if text.len() < 126 {
        frame.push(0x80 | u8::try_from(text.len()).expect("short text"));
    } else if let Ok(length) = u16::try_from(text.len()) {
        frame.push(0x80 | 0x7e);
        frame.extend_from_slice(&length.to_be_bytes());
    } else {
        frame.push(0x80 | 127);
        frame.extend_from_slice(&(text.len() as u64).to_be_bytes());
    }
    frame.extend_from_slice(&mask);
    for (index, byte) in text.as_bytes().iter().enumerate() {
        frame.push(byte ^ mask[index % 4]);
    }
    frame
}
