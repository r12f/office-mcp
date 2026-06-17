use base64::Engine;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::io::Read;
use std::net::{IpAddr, ToSocketAddrs};

const MAX_IMAGE_BYTES: usize = 10 * 1024 * 1024;
const MAX_REDIRECTS: usize = 3;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FetchedImage {
    pub base64: String,
    pub mime_type: ImageMimeType,
    pub byte_length: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageMimeType {
    Png,
    Jpeg,
}

impl ImageMimeType {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageFetcher {
    max_image_bytes: usize,
    max_redirects: usize,
}

impl ImageFetcher {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_image_bytes: MAX_IMAGE_BYTES,
            max_redirects: MAX_REDIRECTS,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "fetches public HTTPS images with SSRF and payload validation"
    }

    /// Validates existing base64 image input and returns normalized metadata.
    ///
    /// # Errors
    ///
    /// Returns an error when base64 decoding fails, the image exceeds the size
    /// limit, or the decoded bytes are not an allowed image format.
    pub fn validate_base64(&self, base64: &str) -> Result<FetchedImage, ImageFetchError> {
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(base64)
            .map_err(|_| ImageFetchError::InvalidImage("Image base64 is invalid.".to_string()))?;
        self.encode_validated_image(&bytes)
    }

    /// Fetches and validates an HTTPS image URL.
    ///
    /// # Errors
    ///
    /// Returns an error when the URL, DNS resolution, HTTP response, redirect,
    /// image size, or image format violates the configured policy.
    pub fn fetch_url(&self, url: &str) -> Result<FetchedImage, ImageFetchError> {
        self.fetch_url_with_redirects(url, 0)
    }

    fn fetch_url_with_redirects(
        &self,
        url: &str,
        redirect_count: usize,
    ) -> Result<FetchedImage, ImageFetchError> {
        let parsed = ParsedHttpsUrl::parse(url)?;
        parsed.assert_public_addresses()?;
        let response = ureq::get(url)
            .set("accept", "image/png,image/jpeg")
            .call()
            .map_err(ImageFetchError::from_ureq)?;
        if is_redirect(response.status()) {
            if redirect_count >= self.max_redirects {
                return Err(ImageFetchError::Policy(
                    "Too many redirects while fetching image.".to_string(),
                ));
            }
            let location = response.header("location").ok_or_else(|| {
                ImageFetchError::Policy("Redirect response is missing Location header.".to_string())
            })?;
            let next_url = resolve_redirect(url, location)?;
            return self.fetch_url_with_redirects(&next_url, redirect_count + 1);
        }
        if !(200..=299).contains(&response.status()) {
            return Err(ImageFetchError::Http(response.status()));
        }
        let mut bytes = Vec::new();
        response
            .into_reader()
            .take(u64::try_from(self.max_image_bytes + 1).unwrap_or(u64::MAX))
            .read_to_end(&mut bytes)
            .map_err(ImageFetchError::Io)?;
        self.encode_validated_image(&bytes)
    }

    fn encode_validated_image(&self, bytes: &[u8]) -> Result<FetchedImage, ImageFetchError> {
        if bytes.len() > self.max_image_bytes {
            return Err(ImageFetchError::InvalidImage(
                "Image exceeds 10 MiB limit.".to_string(),
            ));
        }
        let mime_type = detect_image_mime(bytes).ok_or_else(|| {
            ImageFetchError::InvalidImage("Only PNG and JPEG images are supported.".to_string())
        })?;
        Ok(FetchedImage {
            base64: base64::engine::general_purpose::STANDARD.encode(bytes),
            mime_type,
            byte_length: bytes.len(),
        })
    }
}

impl Default for ImageFetcher {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug)]
pub enum ImageFetchError {
    Policy(String),
    InvalidImage(String),
    Http(u16),
    Io(std::io::Error),
}

impl ImageFetchError {
    fn from_ureq(error: ureq::Error) -> Self {
        match error {
            ureq::Error::Status(status, _) => Self::Http(status),
            ureq::Error::Transport(error) => Self::Policy(error.to_string()),
        }
    }
}

impl Display for ImageFetchError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Policy(message) | Self::InvalidImage(message) => formatter.write_str(message),
            Self::Http(status) => write!(formatter, "Image fetch failed with HTTP {status}."),
            Self::Io(error) => write!(formatter, "{error}"),
        }
    }
}

impl Error for ImageFetchError {}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ParsedHttpsUrl {
    host: String,
    port: u16,
}

impl ParsedHttpsUrl {
    fn parse(url: &str) -> Result<Self, ImageFetchError> {
        let remainder = url
            .strip_prefix("https://")
            .ok_or_else(|| ImageFetchError::Policy("Image URL must use https.".to_string()))?;
        let authority = remainder
            .split(['/', '?', '#'])
            .next()
            .filter(|value| !value.is_empty())
            .ok_or_else(|| {
                ImageFetchError::Policy("Image URL must include a hostname.".to_string())
            })?;
        if authority.contains('@') {
            return Err(ImageFetchError::Policy(
                "Image URL credentials are not allowed.".to_string(),
            ));
        }
        let (host, port) = parse_authority(authority)?;
        Ok(Self { host, port })
    }

    fn assert_public_addresses(&self) -> Result<(), ImageFetchError> {
        let addresses = (self.host.as_str(), self.port)
            .to_socket_addrs()
            .map_err(|error| ImageFetchError::Policy(error.to_string()))?
            .collect::<Vec<_>>();
        if addresses.is_empty() {
            return Err(ImageFetchError::Policy(
                "Image URL hostname did not resolve.".to_string(),
            ));
        }
        for address in addresses {
            if is_forbidden_address(&address.ip()) {
                return Err(ImageFetchError::Policy(format!(
                    "Image URL resolves to a private or reserved address: {}.",
                    address.ip()
                )));
            }
        }
        Ok(())
    }
}

#[must_use]
pub fn is_forbidden_address(address: &IpAddr) -> bool {
    match address {
        IpAddr::V4(address) => {
            let octets = address.octets();
            let [a, b, _, _] = octets;
            a == 0
                || a == 10
                || a == 127
                || (a == 100 && (64..=127).contains(&b))
                || (a == 169 && b == 254)
                || (a == 172 && (16..=31).contains(&b))
                || (a == 192 && b == 168)
                || (a == 198 && (b == 18 || b == 19))
                || a >= 224
        }
        IpAddr::V6(address) => {
            address.is_unspecified()
                || address.is_loopback()
                || address.is_unique_local()
                || address.is_unicast_link_local()
                || address.is_multicast()
        }
    }
}

fn is_redirect(status: u16) -> bool {
    matches!(status, 301 | 302 | 303 | 307 | 308)
}

fn resolve_redirect(base: &str, location: &str) -> Result<String, ImageFetchError> {
    if location.starts_with("https://") || location.starts_with("http://") {
        return Ok(location.to_string());
    }
    let remainder = base
        .strip_prefix("https://")
        .ok_or_else(|| ImageFetchError::Policy("Image URL is invalid.".to_string()))?;
    let (authority, base_path) = remainder
        .split_once('/')
        .map_or((remainder, "/"), |(authority, path)| (authority, path));
    let path = if location.starts_with('/') {
        location.to_string()
    } else {
        let parent = base_path.rsplit_once('/').map_or("/", |(parent, _)| parent);
        format!("{parent}/{location}")
    };
    Ok(format!("https://{authority}{path}"))
}

fn parse_authority(authority: &str) -> Result<(String, u16), ImageFetchError> {
    if let Some(inner) = authority.strip_prefix('[') {
        let (host, rest) = inner
            .split_once(']')
            .ok_or_else(|| ImageFetchError::Policy("Image URL host is invalid.".to_string()))?;
        let port = rest.strip_prefix(':').map_or(Ok(443), parse_port)?;
        return Ok((host.to_string(), port));
    }
    let Some((host, port)) = authority.rsplit_once(':') else {
        return Ok((authority.to_string(), 443));
    };
    if host.contains(':') {
        return Ok((authority.to_string(), 443));
    }
    Ok((host.to_string(), parse_port(port)?))
}

fn parse_port(port: &str) -> Result<u16, ImageFetchError> {
    port.parse::<u16>()
        .map_err(|_| ImageFetchError::Policy("Image URL port is invalid.".to_string()))
}

fn detect_image_mime(bytes: &[u8]) -> Option<ImageMimeType> {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]) {
        return Some(ImageMimeType::Png);
    }
    if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        return Some(ImageMimeType::Jpeg);
    }
    None
}

#[cfg(test)]
#[path = "image_fetcher_tests.rs"]
mod image_fetcher_tests;
