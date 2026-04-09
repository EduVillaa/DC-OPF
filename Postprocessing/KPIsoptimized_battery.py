import pandas as pd
import numpy as np
import pandas as pd


def get_e_nom(grid):
    stores = grid.stores.loc[grid.stores.index.str.startswith("BatteryStore_")]

    if "e_nom_opt" in stores.columns:
        return stores["e_nom_opt"].fillna(stores["e_nom"])
    else:
        return stores["e_nom"]

def get_p_nom(grid):
    links = grid.links.loc[
        grid.links.index.str.startswith("BatteryDischarge_")
    ]

    if "p_nom_opt" in links.columns:
        return links["p_nom_opt"].fillna(links["p_nom"])
    else:
        return links["p_nom"]

def get_snapshot_hours(grid):
    """
    Devuelve una serie con la duración de cada snapshot en horas.
    Si existe snapshot_weightings, usa la columna 'objective'.
    """
    if hasattr(grid, "snapshot_weightings") and "objective" in grid.snapshot_weightings.columns:
        return grid.snapshot_weightings["objective"]
    else:
        return pd.Series(1.0, index=grid.snapshots)


import pandas as pd
import numpy as np


def get_battery_sizes(grid):
    """
    Devuelve un DataFrame con KPIs de las baterías:
    - Energy (MWh)
    - Power (MW)
    - Duration (h)
    - Battery charge (MWh)
    - Battery discharge (MWh)
    - Throughput (MWh)
    - Equivalent cycles
    - Real efficiency
    - Utilization factor
    - Hours SOC 0–5%
    - Hours SOC 95–100%

    Si no hay baterías en la red, devuelve un DataFrame vacío
    con las columnas esperadas.
    """

    output_columns = [
        "Energy (MWh)",
        "Power (MW)",
        "Duration (h)",
        "Battery charge (MWh)",
        "Battery discharge (MWh)",
        "Throughput (MWh)",
        "Equivalent cycles",
        "Real efficiency",
        "Utilization factor",
        "Hours SOC 0–5%",
        "Hours SOC 95–100%",
    ]

    empty_result = pd.DataFrame(columns=output_columns)
    empty_result.index.name = "name"

    # ---------------------------------
    # Comprobaciones básicas de existencia
    # ---------------------------------
    if not hasattr(grid, "stores") or grid.stores is None or grid.stores.empty:
        return empty_result

    if not hasattr(grid, "links") or grid.links is None:
        return empty_result

    # ---------------------------------
    # Identificar stores de batería
    # ---------------------------------
    store_index = pd.Index(grid.stores.index)
    store_names = [str(name) for name in store_index if str(name).startswith("BatteryStore_")]

    if not store_names:
        return empty_result

    # ---------------------------------
    # Capacidades de energía (e_nom)
    # ---------------------------------
    # Si existe e_nom_opt y tiene valor, usarlo; si no, e_nom
    e_nom = pd.Series(index=store_names, dtype=float)

    for store_name in store_names:
        energy_mwh = np.nan

        if "e_nom_opt" in grid.stores.columns:
            val = grid.stores.loc[store_name, "e_nom_opt"]
            if pd.notna(val) and float(val) > 0:
                energy_mwh = float(val)

        if (pd.isna(energy_mwh) or energy_mwh <= 0) and "e_nom" in grid.stores.columns:
            val = grid.stores.loc[store_name, "e_nom"]
            if pd.notna(val):
                energy_mwh = float(val)

        e_nom.loc[store_name] = energy_mwh

    # Si por lo que sea no se pudo leer nada útil
    if e_nom.empty:
        return empty_result

    # ---------------------------------
    # Capacidades de potencia (p_nom)
    # ---------------------------------
    p_nom = pd.Series(dtype=float)

    if hasattr(grid, "links") and grid.links is not None and not grid.links.empty:
        link_names = [str(name) for name in grid.links.index]

        for link_name in link_names:
            power_mw = np.nan

            if "p_nom_opt" in grid.links.columns:
                val = grid.links.loc[link_name, "p_nom_opt"]
                if pd.notna(val) and float(val) > 0:
                    power_mw = float(val)

            if (pd.isna(power_mw) or power_mw <= 0) and "p_nom" in grid.links.columns:
                val = grid.links.loc[link_name, "p_nom"]
                if pd.notna(val):
                    power_mw = float(val)

            p_nom.loc[link_name] = power_mw

    # ---------------------------------
    # Duración de snapshots
    # ---------------------------------
    if hasattr(grid, "snapshot_weightings") and grid.snapshot_weightings is not None:
        if hasattr(grid.snapshot_weightings, "objective"):
            snapshot_hours = grid.snapshot_weightings.objective.copy()
        elif "objective" in grid.snapshot_weightings.columns:
            snapshot_hours = grid.snapshot_weightings["objective"].copy()
        else:
            snapshot_hours = pd.Series(1.0, index=grid.snapshots)
    else:
        snapshot_hours = pd.Series(1.0, index=grid.snapshots)

    snapshot_hours = pd.Series(snapshot_hours, index=grid.snapshots).fillna(1.0)

    # ---------------------------------
    # SOC (%) usando stores_t.e
    # ---------------------------------
    hours_empty = pd.Series(dtype=float)
    hours_full = pd.Series(dtype=float)

    if (
        hasattr(grid, "stores_t")
        and hasattr(grid.stores_t, "e")
        and grid.stores_t.e is not None
        and not grid.stores_t.e.empty
    ):
        store_energy = grid.stores_t.e.copy()

        valid_store_names = [name for name in store_names if name in store_energy.columns and name in e_nom.index]

        if valid_store_names:
            denom = e_nom.loc[valid_store_names].replace(0, np.nan)
            soc_percent = store_energy[valid_store_names].divide(denom, axis=1) * 100

            LOW_THRESHOLD = 5
            HIGH_THRESHOLD = 95

            hours_empty = (soc_percent <= LOW_THRESHOLD).mul(snapshot_hours, axis=0).sum()
            hours_full = (soc_percent >= HIGH_THRESHOLD).mul(snapshot_hours, axis=0).sum()

    # ---------------------------------
    # Series temporales de links
    # ---------------------------------
    links_t_p0 = pd.DataFrame(index=grid.snapshots)
    links_t_p1 = pd.DataFrame(index=grid.snapshots)

    if hasattr(grid, "links_t"):
        if hasattr(grid.links_t, "p0") and grid.links_t.p0 is not None and not grid.links_t.p0.empty:
            links_t_p0 = grid.links_t.p0.copy()
        if hasattr(grid.links_t, "p1") and grid.links_t.p1 is not None and not grid.links_t.p1.empty:
            links_t_p1 = grid.links_t.p1.copy()

    # ---------------------------------
    # Construcción de KPIs
    # ---------------------------------
    rows = []

    for store_name in store_names:
        suffix = store_name.replace("BatteryStore_", "")
        charge_link = f"BatteryCharge_{suffix}"
        discharge_link = f"BatteryDischarge_{suffix}"

        energy_mwh = float(e_nom.get(store_name, np.nan))

        if discharge_link in p_nom.index and pd.notna(p_nom.loc[discharge_link]):
            power_mw = float(p_nom.loc[discharge_link])
        elif charge_link in p_nom.index and pd.notna(p_nom.loc[charge_link]):
            power_mw = float(p_nom.loc[charge_link])
        else:
            power_mw = np.nan

        # Energía de carga y descarga en lado AC
        if charge_link in links_t_p0.columns:
            charge_series = links_t_p0[charge_link].clip(lower=0)
            charge_mwh = float((charge_series * snapshot_hours).sum())
        else:
            charge_mwh = 0.0

        if discharge_link in links_t_p1.columns:
            discharge_series = (-links_t_p1[discharge_link]).clip(lower=0)
            discharge_mwh = float((discharge_series * snapshot_hours).sum())
        else:
            discharge_mwh = 0.0

        throughput_mwh = charge_mwh + discharge_mwh

        if pd.notna(energy_mwh) and energy_mwh > 0:
            equivalent_cycles = throughput_mwh / (2 * energy_mwh)
        else:
            equivalent_cycles = np.nan

        if charge_mwh > 0:
            real_efficiency = discharge_mwh / charge_mwh
        else:
            real_efficiency = np.nan

        total_hours = float(snapshot_hours.sum())
        if pd.notna(power_mw) and power_mw > 0 and total_hours > 0:
            utilization_factor = discharge_mwh / (power_mw * total_hours)
            duration_h = energy_mwh / power_mw if pd.notna(energy_mwh) else np.nan
        else:
            utilization_factor = np.nan
            duration_h = np.nan

        h_empty = float(hours_empty.get(store_name, np.nan))
        h_full = float(hours_full.get(store_name, np.nan))

        rows.append(
            {
                "name": store_name,
                "Energy (MWh)": energy_mwh,
                "Power (MW)": power_mw,
                "Duration (h)": duration_h,
                "Battery charge (MWh)": round(charge_mwh, 1),
                "Battery discharge (MWh)": round(discharge_mwh, 1),
                "Throughput (MWh)": round(throughput_mwh, 1),
                "Equivalent cycles": round(equivalent_cycles, 1) if pd.notna(equivalent_cycles) else np.nan,
                "Real efficiency": round(real_efficiency, 4) if pd.notna(real_efficiency) else np.nan,
                "Utilization factor": round(utilization_factor, 4) if pd.notna(utilization_factor) else np.nan,
                "Hours SOC 0–5%": h_empty,
                "Hours SOC 95–100%": h_full,
            }
        )

    if not rows:
        return empty_result

    df = pd.DataFrame(rows).set_index("name")
    return df