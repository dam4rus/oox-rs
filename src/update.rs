pub trait Update {
    fn update_with(self, other: Self) -> Self;
}

pub fn update_options<T: Update>(lhs: Option<T>, rhs: Option<T>) -> Option<T> {
    match (lhs, rhs) {
        (Some(lhs), Some(rhs)) => Some(lhs.update_with(rhs)),
        (lhs, rhs) => rhs.or(lhs),
    }
}
