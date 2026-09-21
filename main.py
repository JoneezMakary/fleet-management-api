# ============================================================
# FLEET MANAGEMENT - FASTAPI
# Colab Ready
# Interactive API for Frontend / Filters
# ============================================================

# =========================
# 1. Install packages
# =========================



# =========================
# 2. Imports
# =========================

import os
import threading
import pandas as pd
import numpy as np

from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
import uvicorn



# =========================
# 3. Find Excel file
# =========================

EXCEL_PATH = "Fleet_Data_Analysis_Final.xlsx"

if not os.path.exists(EXCEL_PATH):

    print("Fleet_Data_Analysis_Final.xlsx not found.")
    print("Please upload it.")




    EXCEL_PATH = os.path.join(
        "/content",
        uploaded_name
    )


print("Using file:")
print(EXCEL_PATH)


# =========================
# 4. Load Excel sheets
# =========================

excel = pd.ExcelFile(EXCEL_PATH)

print("\nAvailable sheets:")
for sheet in excel.sheet_names:
    print("-", sheet)


def load_sheet(sheet_name):

    if sheet_name in excel.sheet_names:
        return pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet_name
        )

    return pd.DataFrame()


df = load_sheet("final_analytics_data")

booking_demand = load_sheet("booking_demand")
vehicle_utilization = load_sheet("vehicle_utilization")
utilization_by_type = load_sheet("utilization_by_type")
fleet_utilization_summary = load_sheet(
    "fleet_utilization_summary"
)
service_due_summary = load_sheet(
    "service_due_summary"
)
vehicles_due = load_sheet("vehicles_due")
fuel_by_vehicle_type = load_sheet(
    "fuel_by_vehicle_type"
)
vehicle_performance = load_sheet(
    "vehicle_performance"
)
high_consumption_trips = load_sheet(
    "high_consumption_trips"
)
ac_impact = load_sheet("ac_impact")
traffic_impact = load_sheet("traffic_impact")
daily_trend = load_sheet("daily_trend")


# =========================
# 5. Clean dates
# =========================

date_columns = [
    "request_timestamp",
    "approval_timestamp",
    "trip_start_timestamp",
    "trip_end_timestamp"
]

for col in date_columns:

    if col in df.columns:

        df[col] = pd.to_datetime(
            df[col],
            errors="coerce"
        )


# =========================
# 6. Helper Functions
# =========================

def clean_for_json(dataframe):

    if dataframe is None:
        return []

    if dataframe.empty:
        return []

    result = dataframe.copy()

    # Convert datetime columns
    for col in result.columns:

        if pd.api.types.is_datetime64_any_dtype(
            result[col]
        ):

            result[col] = result[col].astype(str)

    # Replace NaN / Infinity
    result = result.replace(
        [np.inf, -np.inf],
        np.nan
    )

    result = result.where(
        pd.notna(result),
        None
    )

    return result.to_dict(
        orient="records"
    )


def apply_filters(
    data,
    vehicle_type=None,
    fuel_type=None,
    traffic_band=None,
    ac_used=None,
    service_status=None,
    start_date=None,
    end_date=None
):

    filtered = data.copy()

    if vehicle_type:
        if "vehicle_type" in filtered.columns:
            filtered = filtered[
                filtered["vehicle_type"]
                .astype(str)
                .str.lower()
                ==
                vehicle_type.lower()
            ]

    if fuel_type:
        if "fuel_type" in filtered.columns:
            filtered = filtered[
                filtered["fuel_type"]
                .astype(str)
                .str.lower()
                ==
                fuel_type.lower()
            ]

    if traffic_band:
        if "traffic_band" in filtered.columns:
            filtered = filtered[
                filtered["traffic_band"]
                .astype(str)
                .str.lower()
                ==
                traffic_band.lower()
            ]

    if ac_used is not None:

        if "ac_used" in filtered.columns:

            # Accept true / false / 1 / 0
            value = str(ac_used).lower()

            if value in [
                "true",
                "1",
                "yes"
            ]:

                filtered = filtered[
                    filtered["ac_used"]
                    .astype(str)
                    .str.lower()
                    .isin(
                        ["true", "1", "yes"]
                    )
                ]

            elif value in [
                "false",
                "0",
                "no"
            ]:

                filtered = filtered[
                    filtered["ac_used"]
                    .astype(str)
                    .str.lower()
                    .isin(
                        ["false", "0", "no"]
                    )
                ]

    if service_status:

        if "service_status" in filtered.columns:

            filtered = filtered[
                filtered["service_status"]
                .astype(str)
                .str.lower()
                ==
                service_status.lower()
            ]

    if start_date:

        if "trip_start_timestamp" in filtered.columns:

            start = pd.to_datetime(
                start_date,
                errors="coerce"
            )

            if pd.notna(start):

                filtered = filtered[
                    filtered[
                        "trip_start_timestamp"
                    ] >= start
                ]

    if end_date:

        if "trip_start_timestamp" in filtered.columns:

            end = pd.to_datetime(
                end_date,
                errors="coerce"
            )

            if pd.notna(end):

                # Include entire end date
                end = end + pd.Timedelta(days=1)

                filtered = filtered[
                    filtered[
                        "trip_start_timestamp"
                    ] < end
                ]

    return filtered


# =========================
# 7. Create FastAPI
# =========================

app = FastAPI(
    title="Fleet Management API",
    description=(
        "Fleet Management Data Analysis API "
        "for Dashboard and Filters"
    ),
    version="1.0.0"
)


# =========================
# 8. CORS
# =========================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# 9. ROOT
# ============================================================

@app.get("/")
def root():

    return {
        "project": "Fleet Management",
        "api": "Fleet Management REST API",
        "version": "1.0.0",
        "status": "running",
        "documentation": "/docs"
    }


# ============================================================
# 10. API INFO
# ============================================================

@app.get("/api")
def api_info():

    return {
        "message": "Fleet Management API is running",
        "version": "1.0.0",
        "endpoints": [
            "/api/kpis",
            "/api/filters",
            "/api/vehicles",
            "/api/booking-demand",
            "/api/fuel-analysis",
            "/api/utilization",
            "/api/service-due",
            "/api/high-consumption",
            "/api/ac-impact",
            "/api/traffic-impact",
            "/api/daily-trend"
        ]
    }


# ============================================================
# 11. FILTER OPTIONS
# ============================================================

@app.get("/api/filters")
def get_filters():

    result = {}

    if "vehicle_type" in df.columns:

        result["vehicle_type"] = sorted(
            df["vehicle_type"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if "fuel_type" in df.columns:

        result["fuel_type"] = sorted(
            df["fuel_type"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if "traffic_band" in df.columns:

        result["traffic_band"] = sorted(
            df["traffic_band"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if "ac_used" in df.columns:

        result["ac_used"] = sorted(
            df["ac_used"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if "service_status" in df.columns:

        result["service_status"] = sorted(
            df["service_status"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if "trip_start_timestamp" in df.columns:

        result["date_range"] = {
            "min": str(
                df["trip_start_timestamp"]
                .min()
                .date()
            ),
            "max": str(
                df["trip_start_timestamp"]
                .max()
                .date()
            )
        }

    return result


# ============================================================
# 12. CORE KPIs
# ============================================================

@app.get("/api/kpis")
def get_kpis(

    vehicle_type: str = None,

    fuel_type: str = None,

    traffic_band: str = None,

    ac_used: str = None,

    service_status: str = None,

    start_date: str = None,

    end_date: str = None
):

    filtered = apply_filters(
        df,
        vehicle_type,
        fuel_type,
        traffic_band,
        ac_used,
        service_status,
        start_date,
        end_date
    )

    total_reservations = len(filtered)

    completed = (
        filtered["status"]
        .astype(str)
        .str.lower()
        .eq("completed")
        .sum()
        if "status" in filtered.columns
        else 0
    )

    actual_rows = (
        filtered["actual_fuel_liters"]
        .notna()
        .sum()
        if "actual_fuel_liters" in filtered.columns
        else 0
    )

    total_distance = (
        filtered["route_km"].sum()
        if "route_km" in filtered.columns
        else 0
    )

    baseline_fuel = (
        filtered[
            "baseline_estimated_fuel_liters"
        ].sum()
        if "baseline_estimated_fuel_liters"
        in filtered.columns
        else 0
    )

    actual_fuel = (
        filtered["actual_fuel_liters"].sum()
        if "actual_fuel_liters" in filtered.columns
        else 0
    )

    variance = (
        filtered["fuel_variance_liters"].sum()
        if "fuel_variance_liters" in filtered.columns
        else 0
    )

    baseline_cost = (
        filtered["estimated_fuel_cost_egp"].sum()
        if "estimated_fuel_cost_egp"
        in filtered.columns
        else 0
    )

    actual_cost = (
        filtered["actual_fuel_cost_egp"].sum()
        if "actual_fuel_cost_egp"
        in filtered.columns
        else 0
    )

    waste_cost = (
        filtered["fuel_waste_cost_egp"].sum()
        if "fuel_waste_cost_egp"
        in filtered.columns
        else 0
    )

    high_consumption = (
        filtered["consumption_status"]
        .eq("High Consumption")
        .sum()
        if "consumption_status"
        in filtered.columns
        else 0
    )

    return {
        "total_reservations": int(
            total_reservations
        ),

        "completed_reservations": int(
            completed
        ),

        "trips_with_actual_fuel": int(
            actual_rows
        ),

        "missing_actual_fuel": int(
            total_reservations - actual_rows
        ),

        "total_distance_km": round(
            float(total_distance),
            2
        ),

        "baseline_fuel_liters": round(
            float(baseline_fuel),
            2
        ),

        "actual_fuel_liters": round(
            float(actual_fuel),
            2
        ),

        "fuel_variance_liters": round(
            float(variance),
            2
        ),

        "baseline_fuel_cost_egp": round(
            float(baseline_cost),
            2
        ),

        "actual_fuel_cost_egp": round(
            float(actual_cost),
            2
        ),

        "fuel_waste_cost_egp": round(
            float(waste_cost),
            2
        ),

        "high_consumption_trips": int(
            high_consumption
        )
    }


# ============================================================
# 13. VEHICLES
# ============================================================

@app.get("/api/vehicles")
def get_vehicles(

    vehicle_type: str = None,

    fuel_type: str = None,

    service_status: str = None
):

    filtered = apply_filters(
        df,
        vehicle_type=vehicle_type,
        fuel_type=fuel_type,
        service_status=service_status
    )

    columns = [
        "vehicle_id",
        "vehicle_type",
        "make",
        "model",
        "vehicle_year",
        "seats",
        "fuel_type",
        "service_status"
    ]

    columns = [
        c for c in columns
        if c in filtered.columns
    ]

    vehicles = (
        filtered[columns]
        .drop_duplicates()
    )

    return {
        "count": len(vehicles),
        "data": clean_for_json(vehicles)
    }


# ============================================================
# 14. BOOKING DEMAND
# ============================================================

@app.get("/api/booking-demand")
def get_booking_demand(

    vehicle_type: str = None,

    start_date: str = None,

    end_date: str = None
):

    filtered = apply_filters(
        df,
        vehicle_type=vehicle_type,
        start_date=start_date,
        end_date=end_date
    )

    if filtered.empty:

        return {
            "count": 0,
            "data": []
        }

    result = (
        filtered
        .groupby("vehicle_type")
        .agg(
            total_reservations=(
                "reservation_id",
                "count"
            ),

            total_distance_km=(
                "route_km",
                "sum"
            ),

            average_distance_km=(
                "route_km",
                "mean"
            ),

            total_passengers=(
                "passengers",
                "sum"
            )
        )
        .reset_index()
    )

    result["booking_share_pct"] = (
        result["total_reservations"]
        /
        result["total_reservations"].sum()
        *
        100
    )

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 15. FUEL ANALYSIS
# ============================================================

@app.get("/api/fuel-analysis")
def get_fuel_analysis(

    vehicle_type: str = None,

    fuel_type: str = None,

    traffic_band: str = None,

    ac_used: str = None
):

    filtered = apply_filters(
        df,
        vehicle_type,
        fuel_type,
        traffic_band,
        ac_used
    )

    filtered = filtered[
        filtered["actual_fuel_liters"]
        .notna()
    ]

    if filtered.empty:

        return {
            "count": 0,
            "data": []
        }

    result = (
        filtered
        .groupby("vehicle_type")
        .agg(
            trips=(
                "reservation_id",
                "count"
            ),

            total_distance_km=(
                "route_km",
                "sum"
            ),

            baseline_fuel_liters=(
                "baseline_estimated_fuel_liters",
                "sum"
            ),

            actual_fuel_liters=(
                "actual_fuel_liters",
                "sum"
            ),

            variance_liters=(
                "fuel_variance_liters",
                "sum"
            ),

            average_variance_pct=(
                "fuel_variance_pct",
                "mean"
            ),

            waste_cost_egp=(
                "fuel_waste_cost_egp",
                "sum"
            )
        )
        .reset_index()
    )

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 16. UTILIZATION
# ============================================================

@app.get("/api/utilization")
def get_utilization(

    vehicle_type: str = None
):

    result = vehicle_utilization.copy()

    if vehicle_type:

        result = result[
            result["vehicle_type"]
            .astype(str)
            .str.lower()
            ==
            vehicle_type.lower()
        ]

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 17. SERVICE DUE
# ============================================================

@app.get("/api/service-due")
def get_service_due():

    return {
        "vehicles_due": len(
            vehicles_due
        ),

        "data": clean_for_json(
            vehicles_due
        )
    }


# ============================================================
# 18. HIGH CONSUMPTION
# ============================================================

@app.get("/api/high-consumption")
def get_high_consumption(

    vehicle_type: str = None,

    traffic_band: str = None,

    ac_used: str = None,

    limit: int = Query(
        100,
        ge=1,
        le=500
    )
):

    filtered = apply_filters(
        df,
        vehicle_type,
        traffic_band=traffic_band,
        ac_used=ac_used
    )

    result = filtered[
        filtered["fuel_variance_pct"] > 15
    ].copy()

    result = result.sort_values(
        "fuel_variance_pct",
        ascending=False
    ).head(limit)

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 19. AC IMPACT
# ============================================================

@app.get("/api/ac-impact")
def get_ac_impact():

    if df.empty:

        return {
            "count": 0,
            "data": []
        }

    result = (
        df[
            df["actual_fuel_liters"]
            .notna()
        ]
        .groupby("ac_used")
        .agg(
            trips=(
                "reservation_id",
                "count"
            ),

            average_variance_pct=(
                "fuel_variance_pct",
                "mean"
            ),

            average_actual_fuel_liters=(
                "actual_fuel_liters",
                "mean"
            ),

            average_baseline_fuel_liters=(
                "baseline_estimated_fuel_liters",
                "mean"
            )
        )
        .reset_index()
    )

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 20. TRAFFIC IMPACT
# ============================================================

@app.get("/api/traffic-impact")
def get_traffic_impact():

    result = (
        df[
            df["actual_fuel_liters"]
            .notna()
        ]
        .groupby("traffic_band")
        .agg(
            trips=(
                "reservation_id",
                "count"
            ),

            average_variance_pct=(
                "fuel_variance_pct",
                "mean"
            ),

            average_actual_fuel_liters=(
                "actual_fuel_liters",
                "mean"
            ),

            average_baseline_fuel_liters=(
                "baseline_estimated_fuel_liters",
                "mean"
            )
        )
        .reset_index()
    )

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
# 21. DAILY TREND
# ============================================================

@app.get("/api/daily-trend")
def get_daily_trend():

    if df.empty:

        return {
            "count": 0,
            "data": []
        }

    temp = df.copy()

    temp["trip_date"] = (
        temp["trip_start_timestamp"]
        .dt.strftime("%Y-%m-%d")
    )

    temp = temp[
        temp["actual_fuel_liters"]
        .notna()
    ]

    result = (
        temp
        .groupby("trip_date")
        .agg(
            trips=(
                "reservation_id",
                "count"
            ),

            total_distance_km=(
                "route_km",
                "sum"
            ),

            baseline_fuel_liters=(
                "baseline_estimated_fuel_liters",
                "sum"
            ),

            actual_fuel_liters=(
                "actual_fuel_liters",
                "sum"
            ),

            variance_liters=(
                "fuel_variance_liters",
                "sum"
            )
        )
        .reset_index()
    )

    result["variance_pct"] = (
        result["variance_liters"]
        /
        result["baseline_fuel_liters"]
        *
        100
    )

    return {
        "count": len(result),
        "data": clean_for_json(result)
    }


# ============================================================
