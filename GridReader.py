import pandas as pd
import pypsa
import matplotlib.pyplot as plt


def excel_reader(filename: str) -> pd.DataFrame:
    df = pd.read_excel(filename, header=8)
    df_settings = pd.read_excel(filename, header=None, nrows=6)
    df["Pmin (MW)"] = pd.to_numeric(df["Pmin (MW)"], errors="coerce").fillna(0)
    df["a (€/MW²h)"] = pd.to_numeric(df["a (€/MW²h)"], errors="coerce").fillna(0)
    df["b (€/MWh)"] = pd.to_numeric(df["b (€/MWh)"], errors="coerce").fillna(0)
    df["c (€)"] = pd.to_numeric(df["c (€)"], errors="coerce").fillna(0)
    df["pwl segments"] = pd.to_numeric(df["pwl segments"], errors="coerce").fillna(1).astype(int)
    df["Loss factor (%)"] = pd.to_numeric(df["Loss factor (%)"], errors="coerce").fillna(0)

    return df, df_settings

def build_network() -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")
    return grid

def add_buses(grid: pypsa.Network, df: pd.DataFrame) -> None:
    n_buses = df["Bus rated voltage (kV)"].count()
    for n in range(n_buses):
        grid.add("Bus", f"Bus_node_{n+1}", v_nom=df.loc[n, "Bus rated voltage (kV)"], carrier="AC")

def add_generators(grid: pypsa.Network, df: pd.DataFrame) -> None:
    for n in range(df["Rated active power (MW)"].last_valid_index() + 1):
        Pmax = df.loc[n, "Rated active power (MW)"]
        if pd.isna(Pmax):
            continue

        Pmin = df.loc[n, "Pmin (MW)"]
        segs = int(df.loc[n, "pwl segments"]) if pd.notna(df.loc[n, "pwl segments"]) else 1

        a = df.loc[n, "a (€/MW²h)"]
        b = df.loc[n, "b (€/MWh)"]

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
                    "Generator", f"Generator_node_{n+1}_seg{i+1}",
                    bus=f"Bus_node_{n+1}",
                    p_nom=step,
                    p_min_pu=p_min_pu,
                    marginal_cost=marginal_cost,
                    carrier="AC"
                )
        else:
            grid.add(
                "Generator", f"Generator_node_{n+1}_seg1",
                bus=f"Bus_node_{n+1}",
                p_nom=Pmax,
                p_min_pu=(Pmin / Pmax) if Pmax > 0 else 0.0,
                marginal_cost=b,
                carrier="AC"
            )

def add_loads(grid: pypsa.Network, df: pd.DataFrame, df_settings: pd.DataFrame) -> None:

    for n in range(df["Active power demand (MW)"].last_valid_index() + 1):
        Pd = df.loc[n, "Active power demand (MW)"]
        Ploss = df.loc[n, "Loss factor (%)"]
        VOLL = df_settings.iat[2, 3] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{n+1}", bus=f"Bus_node_{n+1}", p_set=Pd*(1+Ploss), carrier="AC")

            if df_settings.iat[3, 3] == 1:
                grid.add("Generator", f"shedding_gen_node_{n+1}", bus=f"Bus_node_{n+1}", 
                        p_nom=1e6, 
                        marginal_cost=VOLL,
                        p_min_pu=0,
                        carrier="AC")
    


def add_lines(grid: pypsa.Network, df: pd.DataFrame) -> None:
    for n in range(df["From"].count()):
        desde = int(df.loc[n, "From"])
        hasta = int(df.loc[n, "To"])
        grid.add(
            "Line", f"L{desde}{hasta}",
            bus0=f"Bus_node_{desde}",
            bus1=f"Bus_node_{hasta}",
            x=df.loc[n, "Reactance (p.u)"],
            r=1e-6, #Para evitar el warning que sale al omitir r
            s_nom=df.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )



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
    df,  df_settings= excel_reader("ExampleGrid.xlsx")

    grid = build_network()

    add_buses(grid, df)
    add_generators(grid, df)
    add_loads(grid, df, df_settings)
    add_lines(grid, df)

    solve_opf(grid, solver_name="highs")
    print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])


if __name__ == "__main__":
    main()


