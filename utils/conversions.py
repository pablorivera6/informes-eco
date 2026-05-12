"""Unit conversion helpers for field-reported vs. contract units."""

# Units that require conversion from field ML to M3
M3_KEYWORDS = ("excavaci", "relleno", "subbase", "base granular", "gravas", "material seleccionado",
                "concreto", "demolici", "tipo 4", "tipo 6", "tipo 7", "tipo 10", "tipo 11",
                "mezcla densa", "demolición de pavimento", "grouting", "epoxi")


def needs_conversion(unidad: str, descripcion: str) -> bool:
    """Return True if the item is paid in M3 but field reports ML."""
    if unidad.upper() != "M3":
        return False
    desc_lower = descripcion.lower()
    return any(kw in desc_lower for kw in M3_KEYWORDS)


def default_factor(unidad: str, descripcion: str) -> float:
    """Suggest a default conversion factor (ML → contract unit)."""
    if needs_conversion(unidad, descripcion):
        return 0.30  # typical trench: 0.5m wide × 0.6m deep
    return 1.0


def apply_factor(cantidad_campo: float, factor: float) -> float:
    return round(cantidad_campo * factor, 4)
