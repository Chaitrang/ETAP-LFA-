"""
shortcircuit_fields.py
----------------------
Central catalogue of the columns a Short Circuit report can contain.

Mirrors :mod:`loadflow_fields`: every checkbox the interface offers, and every
column the preview, Word and Excel tables can hold, is one :class:`Field` here.

The catalogue is **data driven**: each field reads a key the Short Circuit
parser actually produces. Fields the ETAP report does not publish at bus level
are marked ``available=False`` so the interface can disable them and say why,
rather than emitting a blank column. Fields that need ETAP's per-bus detail
pages are marked ``needs_detail`` so the expensive scan runs only when they are
selected.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

# --------------------------------------------------------------------------- #
# Units
# --------------------------------------------------------------------------- #

#: The parser reports currents in kA and voltages in kV, as ETAP prints them.
CURRENT_UNITS = {"kA": 1.0, "A": 1000.0}
VOLTAGE_UNITS = {"kV": 1.0, "V": 1000.0}
DEFAULT_CURRENT_UNIT = "kA"
DEFAULT_VOLTAGE_UNIT = "kV"


@dataclass(frozen=True)
class DisplayContext:
    """Unit choices that affect labels and values."""

    current_unit: str = DEFAULT_CURRENT_UNIT
    voltage_unit: str = DEFAULT_VOLTAGE_UNIT

    @property
    def current_scale(self) -> float:
        return CURRENT_UNITS.get(self.current_unit, 1.0)

    @property
    def voltage_scale(self) -> float:
        return VOLTAGE_UNITS.get(self.voltage_unit, 1.0)


INFO = "Info"
RESULTS = "Results"
IDENTITY = "Identity"
STATUS = "Status"


@dataclass(frozen=True)
class Field:
    """One selectable column."""

    key: str
    #: Fixed label, or a callable taking the DisplayContext (for units).
    label: object
    group: str
    getter: Callable[[dict, DisplayContext], object]
    decimals: Optional[int] = 3
    available: bool = True
    note: str = ""
    #: True when the value comes from ETAP's per-bus detail pages.
    needs_detail: bool = False
    trim: bool = False

    def label_for(self, context: DisplayContext) -> str:
        return self.label(context) if callable(self.label) else str(self.label)

    def value(self, row: dict, context: DisplayContext):
        try:
            return self.getter(row, context)
        except Exception:
            return None


# --------------------------------------------------------------------------- #
# Getters
# --------------------------------------------------------------------------- #


def _plain(key: str):
    return lambda row, context: row.get(key)


def _current(key: str):
    """Currents are parsed in kA; scale to the selected unit."""
    def getter(row: dict, context: DisplayContext):
        value = row.get(key)
        return None if value is None else value * context.current_scale
    return getter


def _voltage(key: str):
    def getter(row: dict, context: DisplayContext):
        value = row.get(key)
        return None if value is None else value * context.voltage_scale
    return getter


# --------------------------------------------------------------------------- #
# Catalogue
# --------------------------------------------------------------------------- #

FIELDS: dict[str, Field] = {
    # -- identity (always first, never removable) ---------------------------- #
    "bus_id": Field("bus_id", "ID", IDENTITY, _plain("bus_id"), decimals=None),

    # -- Info --------------------------------------------------------------- #
    "nominal_kv": Field(
        "nominal_kv", lambda c: f"Nominal {c.voltage_unit}", INFO,
        _voltage("nominal_kv"), decimals=3, trim=True,
    ),
    "bus_type": Field(
        "bus_type", "Type", INFO, _plain("bus_type"), decimals=None,
        note="Bus, Switchgear, Switchboard, MCC, Switchrack or Panelboard, "
             "as ETAP classifies the bus in the summary table.",
    ),
    "cfactor": Field(
        "cfactor", "Cfactor", INFO, _plain("cfactor"), decimals=2, needs_detail=True,
        note="IEC voltage factor c, read from each bus's detail page. "
             "Selecting it makes the application scan the whole report.",
    ),
    "rating_peak": Field(
        "rating_peak", lambda c: f"Rated ip ({c.current_unit})", INFO,
        _current("rating_peak"), decimals=2,
        note="Equipment making (peak) capacity. Blank where the report gives "
             "the bus no rated device.",
    ),
    "rating_ib_sym": Field(
        "rating_ib_sym", lambda c: f"Rated Ib sym ({c.current_unit})", INFO,
        _current("rating_ib_sym"), decimals=2,
        note="Equipment breaking capacity; published for only a few devices.",
    ),
    "governing_device": Field(
        "governing_device", "Rated device", INFO,
        _plain("governing_device"), decimals=None,
        note="The device whose rating governs the bus (the lowest rated one).",
    ),
    "xr_ratio": Field(
        "xr_ratio", "X/R", INFO, _plain("xr_ratio"), decimals=1, needs_detail=True,
        note="X/R ratio of the total fault contribution, from each bus's "
             "detail page. Selecting it makes the application scan the whole report.",
    ),

    # -- Results ------------------------------------------------------------ #
    "ik_initial": Field(
        "ik_initial", lambda c: f'I"k ({c.current_unit})', RESULTS,
        _current("ik_initial"), decimals=3,
        note="Initial symmetrical short-circuit current.",
    ),
    "ip_peak": Field(
        "ip_peak", lambda c: f"ip ({c.current_unit})", RESULTS,
        _current("ip_peak"), decimals=3, note="Peak short-circuit current.",
    ),
    "ik_steady": Field(
        "ik_steady", lambda c: f"Ik ({c.current_unit})", RESULTS,
        _current("ik_steady"), decimals=3, note="Steady state short-circuit current.",
    ),
    "ib_sym": Field(
        "ib_sym", lambda c: f"Ib sym ({c.current_unit})", RESULTS,
        _current("ib_sym"), decimals=3, available=False,
        note="The summary table publishes Ib only per protective device, not "
             "per bus; the per-bus figure is time-dependent (one value per "
             "breaking time), so there is no single correct value to report.",
    ),
    "idc": Field(
        "idc", lambda c: f"Idc ({c.current_unit})", RESULTS,
        _current("idc"), decimals=3, available=False,
        note="As Ib sym: published per device and per breaking time only.",
    ),
    "standard": Field(
        "standard", "Standard", RESULTS, _plain("standard"), decimals=None,
        note="The short-circuit standard the study was run under.",
    ),

    # -- status ------------------------------------------------------------- #
    "duty_percent": Field(
        "duty_percent", "Duty (%)", STATUS, _plain("duty_percent"), decimals=1,
        note="Peak duty as a percentage of the equipment rating: ip / Rated ip.",
    ),
    "remarks": Field("remarks", "Remarks", STATUS, _plain("remarks"), decimals=None),
}

#: Group order, and the order of fields inside each group.
INFO_ORDER = (
    "nominal_kv",
    "bus_type",
    "cfactor",
    "rating_peak",
    "rating_ib_sym",
    "governing_device",
    "xr_ratio",
)
RESULT_ORDER = ("ik_initial", "ip_peak", "ik_steady", "ib_sym", "idc", "standard")

#: Defaults follow the reference interface, restricted to fields this report
#: actually publishes. Cfactor is left off by default because selecting it
#: triggers the whole-document detail scan.
DEFAULT_INFO = ("nominal_kv", "rating_peak")
DEFAULT_RESULTS = ("ik_initial", "ip_peak", "ik_steady")

#: Equipment types ETAP uses for a bus row, for the Device Type filter.
EQUIPMENT_TYPES = (
    "Bus",
    "Switchgear",
    "Switchboard",
    "Mcc",
    "Switchrack",
    "Panelboard",
)


def get(key: str) -> Field:
    return FIELDS[key]


def available_keys(order) -> list[str]:
    return [key for key in order if FIELDS[key].available]


def needs_detail(keys) -> bool:
    """True when any selected field requires the per-bus detail pages."""
    return any(FIELDS[key].needs_detail for key in keys if key in FIELDS)
