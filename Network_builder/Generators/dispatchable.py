import pypsa
import pandas as pd


def add_dispatchable_generators(grid: pypsa.Network, df_Gen_Dispatchable: pd.DataFrame) -> None:

    df = df_Gen_Dispatchable.copy()

    df["Rated active power (MW)"] = pd.to_numeric(
        df["Rated active power (MW)"], errors="coerce"
    )
    df["Pmin (MW)"] = pd.to_numeric(
        df["Pmin (MW)"], errors="coerce"
    ).fillna(0.0)

    df["Ramp limit up (p.u)"] = pd.to_numeric(
        df["Ramp limit up (p.u)"], errors="coerce"
    ).fillna(1)
    df["Ramp limit down (p.u)"] = pd.to_numeric(
        df["Ramp limit down (p.u)"], errors="coerce"
    ).fillna(1)
    df["€/MW²h"] = pd.to_numeric(
        df["€/MW²h"], errors="coerce"
    ).fillna(0.0)

    df["€/MWh"] = pd.to_numeric(
        df["€/MWh"], errors="coerce"
    ).fillna(0.0)

    for n in range(df["GENERATOR LOCATION"].count()):
        Pmax = df.loc[n, "Rated active power (MW)"]
        location = df.loc[n, "GENERATOR LOCATION"]

        if pd.isna(Pmax) or pd.isna(location) or Pmax <= 0:
            continue

        Pmax = float(Pmax)
        location = int(location)

        Pmin = float(df.loc[n, "Pmin (MW)"])
        marginal_cost = float(df.loc[n, "€/MWh"])
     
        marginal_cost_quadratic = float(df.loc[n, "€/MW²h"])
        #ramp_limit_up = 1
        #ramp_limit_down = 1

        p_min_pu = Pmin / Pmax if Pmax > 0 else 0.0

        grid.add(
            "Generator",
            f"DispatchGen{location}_{n}",   # n diferencia generadores en el mismo bus
            bus=f"Bus_node_{location}",
            p_nom=Pmax,
            p_min_pu=p_min_pu,
            marginal_cost=marginal_cost,
            marginal_cost_quadratic=marginal_cost_quadratic,
            ramp_limit_up = 0.3,
            ramp_limit_down = 0.3,
            committable =True,
        
            carrier="AC",
        )


