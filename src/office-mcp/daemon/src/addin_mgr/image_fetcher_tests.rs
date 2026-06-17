use super::{ImageFetcher, ImageMimeType, is_forbidden_address};
use base64::Engine;
use std::net::{IpAddr, Ipv4Addr, Ipv6Addr};

const PNG_1X1: &[u8] = &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 0x00];

#[test]
fn rejects_private_and_reserved_ip_addresses() {
    assert!(is_forbidden_address(&IpAddr::V4(Ipv4Addr::LOCALHOST)));
    assert!(is_forbidden_address(&IpAddr::V4(Ipv4Addr::new(
        10, 1, 2, 3
    ))));
    assert!(is_forbidden_address(&IpAddr::V4(Ipv4Addr::new(
        172, 16, 0, 1
    ))));
    assert!(is_forbidden_address(&IpAddr::V4(Ipv4Addr::new(
        192, 168, 1, 1
    ))));
    assert!(is_forbidden_address(&IpAddr::V4(Ipv4Addr::new(
        169, 254, 169, 254
    ))));
    assert!(is_forbidden_address(&IpAddr::V6(Ipv6Addr::LOCALHOST)));
    assert!(!is_forbidden_address(&IpAddr::V4(Ipv4Addr::new(
        8, 8, 8, 8
    ))));
}

#[test]
fn validates_base64_png_images() {
    let base64 = base64::engine::general_purpose::STANDARD.encode(PNG_1X1);
    let image = ImageFetcher::new().validate_base64(&base64).expect("png");

    assert_eq!(image.mime_type, ImageMimeType::Png);
    assert_eq!(image.byte_length, PNG_1X1.len());
    assert_eq!(image.base64, base64);
}

#[test]
fn rejects_non_image_base64_payloads() {
    let base64 = base64::engine::general_purpose::STANDARD.encode([1_u8, 2, 3]);
    let error = ImageFetcher::new()
        .validate_base64(&base64)
        .expect_err("reject non image");

    assert!(error.to_string().contains("PNG and JPEG"));
}

#[test]
fn rejects_non_https_urls_before_fetching() {
    let error = ImageFetcher::new()
        .fetch_url("http://example.com/a.png")
        .expect_err("reject http");

    assert!(error.to_string().contains("https"));
}
