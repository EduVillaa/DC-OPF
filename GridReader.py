import pandas as pd
import pypsa
import matplotlib.pyplot as plt


def leerhojas(filename: str) -> dict:

    sheets = {}

    # --- SYS SETTINGS ---
    sheets["SYS_settings"] = pd.read_excel(
        filename,
        sheet_name="SYS_settings",
        header=1
    )

    # --- NET BUSES ---
    sheets["Net_Buses"] = pd.read_excel(
        filename,
        sheet_name="Net_Buses",
        header=1   # 👈 aquí estaba tu problema
    ).iloc[:, 1:]  # quitar primera columna si hace falta

    # --- NET LINES ---
    sheets["Net_Lines"] = pd.read_excel(
        filename,
        sheet_name="Net_Lines",
        header=2
    ).iloc[:, 1:]

    # --- NET LOADS ---
    sheets["Net_Loads"] = pd.read_excel(
        filename,
        sheet_name="Net_Loads",
        header=2
    ).iloc[:, 1:]

    # --- GENERADORES ---
    sheets["Gen_Dispatchable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Dispatchable",
        header=2
    ).iloc[:, 1:]

    sheets["Gen_Renewable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Renewable",
        header=2
    ).iloc[:, 1:]

    return sheets

def build_network() -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")

    grid.set_snapshots(pd.DatetimeIndex(["2026-01-01 00:00"]))

    return grid

def add_buses(grid: pypsa.Network, df_Net_Buses: pd.DataFrame) -> None:
    n_buses = df_Net_Buses["Bus rated voltage (kV)"].count()
    for n in range(n_buses):
        grid.add("Bus", f"Bus_node_{n+1}", v_nom=df_Net_Buses.loc[n, "Bus rated voltage (kV)"], carrier="AC")

def add_dispatchable_generators(grid: pypsa.Network, df_Gen_Dispatchable: pd.DataFrame) -> None:

    df_Gen_Dispatchable["Pmin (MW)"] = pd.to_numeric(df_Gen_Dispatchable["Pmin (MW)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["a (€/MW²h)"] = pd.to_numeric(df_Gen_Dispatchable["a (€/MW²h)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["b (€/MWh)"] = pd.to_numeric(df_Gen_Dispatchable["b (€/MWh)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["c (€)"] = pd.to_numeric(df_Gen_Dispatchable["c (€)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["pwl segments"] = pd.to_numeric(df_Gen_Dispatchable["pwl segments"], errors="coerce").fillna(1).astype(int)

    for n in range(df_Gen_Dispatchable["Rated active power (MW)"].last_valid_index() + 1):
        Pmax = df_Gen_Dispatchable.loc[n, "Rated active power (MW)"]
        if pd.isna(Pmax):
            continue

        Pmin = df_Gen_Dispatchable.loc[n, "Pmin (MW)"]
        segs = int(df_Gen_Dispatchable.loc[n, "pwl segments"]) if pd.notna(df_Gen_Dispatchable.loc[n, "pwl segments"]) else 1

        a = df_Gen_Dispatchable.loc[n, "a (€/MW²h)"]
        b = df_Gen_Dispatchable.loc[n, "b (€/MWh)"]

        if segs > 1:
            step = Pmax / segs
            remaining_min = Pmin

            for i in range(segs):
                block_min_mw = max(0.0, min(step, remaining_min))
                remaining_min -= block_min_mw

                p_min_pu = block_min_mw / step  # p.u. del bloque

                P_mid = (i + 0.5) * step
                marginal_cost = 2 * a * P_mid + b

                grid.add(
                    "Generator", f"DispatchGen_{n+1}_seg{i+1}",
                    bus=f"Bus_node_{n+1}",
                    p_nom=step,
                    p_min_pu=p_min_pu,
                    marginal_cost=marginal_cost,
                    carrier="AC"
                )
        else:
            grid.add(
                "Generator", f"DispatchGen_{n+1}_seg1",
                bus=f"Bus_node_{n+1}",
                p_nom=Pmax,
                p_min_pu=(Pmin / Pmax) if Pmax > 0 else 0.0,
                marginal_cost=b,
                carrier="AC"
            )

def add_loads(grid: pypsa.Network, df_Net_Loads: pd.DataFrame, df_SYS_settings: pd.DataFrame) -> None:
    df_Net_Loads["Loss factor (%)"] = pd.to_numeric(df_Net_Loads["Loss factor (%)"], errors="coerce").fillna(0)
    for n in range(df_Net_Loads["Active power demand (MW)"].last_valid_index() + 1):
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]
        Ploss = df_Net_Loads.loc[n, "Loss factor (%)"]
        VOLL = df_SYS_settings.iat[0, 2] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{n+1}", bus=f"Bus_node_{n+1}", p_set=Pd*(1+Ploss), carrier="AC")

            use_shed = int(df_SYS_settings.loc[0, "SYSTEM PARAMETERS"]) == 1
            if use_shed:
                grid.add("Generator", f"shedding_gen_node_{n+1}", bus=f"Bus_node_{n+1}", 
                        p_nom=1e6, 
                        marginal_cost=VOLL,
                        p_min_pu=0,
                        carrier="AC")

def add_lines(grid: pypsa.Network, df_Net_Lines: pd.DataFrame) -> None:
    for n in range(df_Net_Lines["From"].count()):
        desde = int(df_Net_Lines.loc[n, "From"])
        hasta = int(df_Net_Lines.loc[n, "To"])
        grid.add(
            "Line", f"L{desde}{hasta}",
            bus0=f"Bus_node_{desde}",
            bus1=f"Bus_node_{hasta}",
            x=df_Net_Lines.loc[n, "Reactance (p.u)"],
            r=1e-6, #Para evitar el warning que sale al omitir r
            s_nom=df_Net_Lines.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )

def add_renewable_generator(grid: pypsa.Network, df_Gen_Dispatchable: pd.DataFrame) -> None:
    for n in range(df_Gen_Dispatchable["Rated active power (MW)"].last_valid_index() + 1):
        if pd.notna(df_Gen_Dispatchable.loc[n, "Rated active power (MW)"]):
            grid.add("Generator", f"RenewableGen_{n+1}", 
                    p_nom= df_Gen_Dispatchable.loc[n, "Rated active power (MW)"],
                    p_min_pu=0,
                    marginal_cost=0, carrier="AC")

def solve_opf(grid: pypsa.Network, solver_name="highs") -> None:
    grid.optimize(solver_name=solver_name)

def export_results(grid: pypsa.Network, filename="results.xlsx") -> None:
    with pd.ExcelWriter(filename) as writer:
        grid.generators_t.p.to_excel(writer, sheet_name="Dispatch")
        grid.lines_t.p0.to_excel(writer, sheet_name="Line_Flows")
        grid.loads_t.p_set.to_excel(writer, sheet_name="Loads")

        cost = (grid.generators_t.p * grid.generators.marginal_cost).sum(axis=1)
        cost.to_frame("Total_Cost").to_excel(writer, sheet_name="Cost")

def main():
    data = leerhojas("ExampleGrid.xlsx")

    df_SYS_settings = data["SYS_settings"]
    df_Net_Buses = data["Net_Buses"]
    df_Net_Lines = data["Net_Lines"]
    df_Net_Loads = data["Net_Loads"]
    df_Gen_Dispatchable = data["Gen_Dispatchable"]
    df_Gen_Renewable = data["Gen_Renewable"]
    
    grid = build_network()

    add_buses(grid, df_Net_Buses)
    add_lines(grid, df_Net_Lines)
    add_loads(grid, df_Net_Loads, df_SYS_settings)
    add_dispatchable_generators(grid, df_Gen_Dispatchable)
    add_renewable_generator(grid, df_Gen_Renewable)

    #solve_opf(grid, solver_name="highs")
    print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])


if __name__ == "__main__":
    main()


