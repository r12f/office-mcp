use super::UiFixtureState;

#[test]
fn fixture_state_defaults_to_seeded_unless_empty_is_requested() {
    assert_eq!(UiFixtureState::from_value(None), UiFixtureState::Seeded);
    assert_eq!(
        UiFixtureState::from_value(Some("seeded")),
        UiFixtureState::Seeded
    );
    assert_eq!(
        UiFixtureState::from_value(Some("empty")),
        UiFixtureState::Empty
    );
    assert_eq!(
        UiFixtureState::from_value(Some("EMPTY")),
        UiFixtureState::Empty
    );
}
