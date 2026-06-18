use super::{ICON_HEIGHT, ICON_WIDTH, product_icon_rgba};
use std::collections::BTreeSet;

#[test]
fn product_tray_icon_is_not_blank_or_single_color() {
    let rgba = product_icon_rgba();

    assert_eq!(rgba.len(), (ICON_WIDTH * ICON_HEIGHT * 4) as usize);
    let colors = rgba
        .chunks_exact(4)
        .map(|chunk| [chunk[0], chunk[1], chunk[2], chunk[3]])
        .collect::<BTreeSet<_>>();

    assert!(
        colors.len() >= 5,
        "icon should contain the full product glyph palette"
    );
    assert!(colors.contains(&[15, 23, 42, 255]));
    assert!(colors.contains(&[72, 214, 164, 255]));
    assert!(colors.contains(&[248, 216, 74, 255]));
    assert!(
        colors.iter().any(|color| color[3] == 0),
        "rounded icon should keep transparent corners"
    );
}

#[test]
fn product_tray_icon_has_visible_control_node() {
    let rgba = product_icon_rgba();
    let yellow_pixels = rgba
        .chunks_exact(4)
        .filter(|chunk| *chunk == [248, 216, 74, 255])
        .count();

    assert!(
        yellow_pixels >= 10,
        "control node should remain visible at tray size"
    );
}
