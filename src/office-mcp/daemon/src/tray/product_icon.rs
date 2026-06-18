pub(crate) const ICON_WIDTH: u32 = 32;
pub(crate) const ICON_HEIGHT: u32 = 32;

#[must_use]
pub(crate) fn product_icon_rgba() -> Vec<u8> {
    let mut rgba = Vec::with_capacity((ICON_WIDTH * ICON_HEIGHT * 4) as usize);
    for y in 0..ICON_HEIGHT {
        for x in 0..ICON_WIDTH {
            rgba.extend_from_slice(&pixel(x, y));
        }
    }
    rgba
}

fn pixel(x: u32, y: u32) -> [u8; 4] {
    let scale = f64::from(ICON_WIDTH) / 256.0;
    let px = (f64::from(x) + 0.5) / scale;
    let py = (f64::from(y) + 0.5) / scale;
    if !rounded_rect(px, py, 0.0, 0.0, 256.0, 256.0, 56.0) {
        return [0, 0, 0, 0];
    }

    let mut color = [15, 23, 42, 255];
    if rounded_rect(px, py, 56.0, 66.0, 118.0, 124.0, 10.0) {
        color = [234, 242, 255, 255];
    }
    if rounded_rect(px, py, 112.0, 54.0, 110.0, 148.0, 10.0) {
        color = [72, 214, 164, 255];
    }
    if line(px, py, 82.0, 92.0, 154.0, 92.0, 10.0)
        || line(px, py, 82.0, 118.0, 140.0, 118.0, 10.0)
        || line(px, py, 82.0, 144.0, 124.0, 144.0, 10.0)
    {
        color = [36, 87, 214, 255];
    }
    if line(px, py, 134.0, 92.0, 192.0, 92.0, 10.0)
        || line(px, py, 134.0, 120.0, 178.0, 120.0, 10.0)
        || line(px, py, 134.0, 148.0, 192.0, 148.0, 10.0)
    {
        color = [15, 23, 42, 184];
    }

    let control_distance = distance(px, py, 190.0, 188.0);
    if control_distance <= 40.0 {
        color = [15, 23, 42, 255];
    }
    if control_distance <= 28.0 {
        color = [248, 216, 74, 255];
    }
    if line(px, py, 178.0, 188.0, 202.0, 188.0, 10.0)
        || line(px, py, 190.0, 176.0, 190.0, 200.0, 10.0)
    {
        color = [15, 23, 42, 255];
    }
    color
}

fn rounded_rect(px: f64, py: f64, x: f64, y: f64, width: f64, height: f64, radius: f64) -> bool {
    if px < x || px > x + width || py < y || py > y + height {
        return false;
    }
    let inner_x = px >= x + radius && px <= x + width - radius;
    let inner_y = py >= y + radius && py <= y + height - radius;
    if inner_x || inner_y {
        return true;
    }
    let rx = px.clamp(x + radius, x + width - radius);
    let ry = py.clamp(y + radius, y + height - radius);
    distance(px, py, rx, ry) <= radius
}

fn line(px: f64, py: f64, x1: f64, y1: f64, x2: f64, y2: f64, width: f64) -> bool {
    let length_sq = (x2 - x1).powi(2) + (y2 - y1).powi(2);
    let t = (((px - x1) * (x2 - x1) + (py - y1) * (y2 - y1)) / length_sq).clamp(0.0, 1.0);
    let ix = x1 + t * (x2 - x1);
    let iy = y1 + t * (y2 - y1);
    distance(px, py, ix, iy) <= width / 2.0
}

fn distance(x1: f64, y1: f64, x2: f64, y2: f64) -> f64 {
    ((x1 - x2).powi(2) + (y1 - y2).powi(2)).sqrt()
}

#[cfg(test)]
#[path = "product_icon_tests.rs"]
mod product_icon_tests;
