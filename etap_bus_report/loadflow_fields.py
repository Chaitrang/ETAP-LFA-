"""
loadflow_fields.py
------------------
Central catalogue of the columns a Load Flow report can contain.

Every column the UI offers - and every column the Word, Excel and preview
tables can hold - is one :class:`Field` here.  Adding a new ETAP quantity later
(power factor, bus angle, generation, frequency ...) is a single entry plus the
value in the parser; no UI, writer or preview code changes.

Fields read from :class:`utils.BusRecord`, which the existing Load Flow parser
already produces - nothing is parsed a second time and no value is invented.
Where ETAP does not publish a quantity the field is marked ``available=False``
so the interface can say so instead of emitting a blank column.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

# --------------------------------------------------------------------------- #
# Display context
# --------------------------------------------------------------------------- #

#: Power units offered in the Display Options group.
POWER_UNITS = {
    "kVA": {"scale": 1.0, "active": "kW", "reactive": "kvar", "apparent": "kVA"},
    "MVA": {"scale": 0.001, "active": "MW", "reactive": "Mvar", "apparent": "MVA"},
}
DEFAULT_POWER_UNIT = "kVA"

#: Voltage display options.
VOLTAGE_DISPLAYS = ("%", "Actual Value")
DEFAULT_VOLTAGE_DISPLAY = "%"


@dataclass(frozen=True)
class DisplayContext:
    """User display choices that affect labels and values."""

    power_unit: str = DEFAULT_POWER_UNIT
    voltage_display: str = DEFAULT_VOLTAGE_DISPLAY

    @property
    def power(self) -> dict:
        return POWER_UNITS.get(self.power_unit, POWER_UNITS[DEFAULT_POWER_UNIT])

    @property
    def scale(self) -> float:
        return self.power["scale"]

    @property
    def voltage_is_actual(self) -> bool:
        return self.voltage_display == "Actual Value"


# --------------------------------------------------------------------------- #
# Field definition
# --------------------------------------------------------------------------- #

BUS_INFO = "Bus Info"
RESULTS = "Load Flow Results"
IDENTITY = "Identity"
STATUS = "Status"


@dataclass(frozen=True)
class Field:
    """One selectable column."""

    key: str
    #: Either a fixed label, or a callable taking the DisplayContext.
    label: object
    group: str
    getter: Callable[[object, DisplayContext], object]
    #: Decimal places; ``None`` renders the value as text.
    decimals: Optional[int] = 1
    #: False when ETAP's Load Flow report simply does not publish the quantity.
    available: bool = True
    #: Shown in the UI when the field is unavailable or derived.
    note: str = ""
    #: Drop trailing zeros (132.000 -> 132, 0.380 -> 0.38), as ETAP's own
    #: results grid does for nominal voltage.
    trim: bool = False

    def label_for(self, context: DisplayContext) -> str:
        return self.label(context) if callable(self.label) else str(self.label)

    def value(self, record, context: DisplayContext):
        try:
            return self.getter(record, context)
        except Exception:
            return None


# --------------------------------------------------------------------------- #
# Getters
# --------------------------------------------------------------------------- #


def _voltage(record, context: DisplayContext):
    """Voltage as ETAP publishes it (%), or converted to actual kV."""
    if record.voltage_percent is None:
        return None
    if context.voltage_is_actual:
        if record.nominal_kv is None:
            return None
        return record.nominal_kv * record.voltage_percent / 100.0
    return record.voltage_percent


def _scaled(meta_key: str):
    def getter(record, context: DisplayContext):
        value = record.meta.get(meta_key)
        return None if value is None else value * context.scale
    return getter


def _kw(record, context: DisplayContext):
    return None if record.kw_loading is None else record.kw_loading * context.scale


def _apparent(record, context: DisplayContext):
    mva = record.meta.get("mva")
    # meta stores MVA; kVA is 1000x that, so the shared scale applies to kVA.
    return None if mva is None else mva * 1000.0 * context.scale


# --------------------------------------------------------------------------- #
# Catalogue
# --------------------------------------------------------------------------- #

FIELDS: dict[str, Field] = {
    # -- identity (always present, never selectable) ------------------------ #
    "bus_id": Field("bus_id", "Bus ID", IDENTITY, lambda r, c: r.bus_id, decimals=None),

    # -- Bus Info ----------------------------------------------------------- #
    "nominal_kv": Field(
        "nominal_kv", "Nominal kV", BUS_INFO,
        lambda r, c: r.nominal_kv, decimals=3, trim=True,
    ),
    "rated_amp": Field(
        "rated_amp", "Amp Rating", BUS_INFO,
        lambda r, c: r.meta.get("rated_amp"), decimals=1,
        note="From the Bus Loading Summary; blank for buses ETAP gives no continuous rating for.",
    ),
    "bus_type": Field(
        "bus_type", "Type", BUS_INFO,
        lambda r, c: None, decimals=None, available=False,
        note="ETAP's Load Flow report does not publish the bus equipment type - "
             "it appears in the Short Circuit report only.",
    ),

    # -- Load Flow Results -------------------------------------------------- #
    "voltage": Field(
        "voltage",
        lambda c: "Voltage (kV)" if c.voltage_is_actual else "Voltage (%)",
        RESULTS, _voltage, decimals=2,
        note="Actual value = Nominal kV x Voltage % / 100.",
    ),
    "kw_loading": Field(
        "kw_loading", lambda c: f"{c.power['active']} Loading", RESULTS, _kw, decimals=1,
    ),
    "kvar_loading": Field(
        "kvar_loading", lambda c: f"{c.power['reactive']} Loading", RESULTS,
        _scaled("kvar_loading"), decimals=1,
        note="Derived from the Bus Loading Summary as S x sin(acos(PF)), "
             "consistent with the active loading column.",
    ),
    "apparent_loading": Field(
        "apparent_loading", lambda c: f"{c.power['apparent']} Loading", RESULTS,
        _apparent, decimals=3,
        note="Total bus load, as ETAP prints it in the Bus Loading Summary.",
    ),
    "amp_loading": Field(
        "amp_loading", "Amp Loading", RESULTS,
        lambda r, c: r.amp_loading, decimals=1,
    ),
    "percent_loading": Field(
        "percent_loading", "% Loading", RESULTS,
        lambda r, c: r.meta.get("percent_loading"), decimals=1,
        note="Amp loading as a percentage of the bus continuous rating; "
             "blank where ETAP gives no rating.",
    ),

    # -- status ------------------------------------------------------------- #
    "alert_cause": Field(
        "alert_cause", "Alert", STATUS, lambda r, c: None, decimals=None,
        note="Which limit triggered the remark (Overvoltage, Undervoltage, Loading).",
    ),
    "remarks": Field("remarks", "Remarks", STATUS, lambda r, c: None, decimals=None),
}

#: Order of the selectable groups, and of the fields inside them.
BUS_INFO_ORDER = ("nominal_kv", "rated_amp", "bus_type")
RESULT_ORDER = (
    "voltage",
    "kw_loading",
    "kvar_loading",
    "apparent_loading",
    "amp_loading",
    "percent_loading",
)

#: What the application produced before this interface existed - used as the
#: default selection so an existing user gets a familiar report with one click.
DEFAULT_BUS_INFO = ("nominal_kv",)
DEFAULT_RESULTS = ("voltage", "kw_loading", "amp_loading")


def get(key: str) -> Field:
    return FIELDS[key]


def available_keys(order) -> list[str]:
    return [key for key in order if FIELDS[key].available]


def unavailable_keys(order) -> list[str]:
    return [key for key in order if not FIELDS[key].available]
