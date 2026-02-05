import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

input_file = r"c:\Users\HP\Desktop\data haraka\Rslt_Mvt_Enseignant_Prim2025_cleaned.xlsx"

print("Loading data...")
df = pd.read_excel(input_file, dtype=str)

df.columns = [col.strip() for col in df.columns]

# Convert numeric columns
if "Points" in df.columns:
    df["Points"] = pd.to_numeric(df["Points"], errors="coerce")
if "Choice_Rank" in df.columns:
    df["Choice_Rank"] = pd.to_numeric(df["Choice_Rank"], errors="coerce")

# Example filter (Rabat)
filtered_data = df[df["Assignment Province"].astype(str).str.contains("الرباط", na=False)]
print(f"Rows for Rabat: {len(filtered_data)}")

# Province analytics
province_stats = (
    df.groupby("Assignment Province", dropna=False)
    .agg(
        total_count=("Assignment Province", "size"),
        avg_points=("Points", "mean"),
        median_points=("Points", "median"),
        rank1_count=("Choice_Rank", lambda s: (s == 1).sum()),
    )
    .reset_index()
)
province_stats["rank1_rate"] = province_stats["rank1_count"] / province_stats["total_count"]

province_stats_sorted = province_stats.sort_values(
    ["avg_points", "rank1_rate"], ascending=[False, False]
)

province_stats_sorted.to_csv(
    r"c:\Users\HP\Desktop\data haraka\province_stats.csv", index=False, encoding="utf-8-sig"
)
province_stats_sorted.to_excel(
    r"c:\Users\HP\Desktop\data haraka\province_stats.xlsx", index=False
)

# Subject analytics
if "Subject" in df.columns:
    subject_stats = (
        df.groupby("Subject", dropna=False)
        .agg(
            total_count=("Subject", "size"),
            avg_points=("Points", "mean"),
            median_points=("Points", "median"),
        )
        .reset_index()
        .sort_values("avg_points", ascending=False)
    )
    subject_stats.to_csv(
        r"c:\Users\HP\Desktop\data haraka\subject_stats.csv", index=False, encoding="utf-8-sig"
    )
    subject_stats.to_excel(
        r"c:\Users\HP\Desktop\data haraka\subject_stats.xlsx", index=False
    )

# Heatmap: top provinces x top subjects by count
if "Assignment Province" in df.columns and "Subject" in df.columns:
    top_provinces = (
        df["Assignment Province"].value_counts().head(15).index.tolist()
    )
    top_subjects = df["Subject"].value_counts().head(15).index.tolist()

    heatmap_df = df[
        df["Assignment Province"].isin(top_provinces)
        & df["Subject"].isin(top_subjects)
    ]

    pivot = pd.pivot_table(
        heatmap_df,
        index="Assignment Province",
        columns="Subject",
        values="Points",
        aggfunc="mean",
    )

    plt.figure(figsize=(14, 8))
    sns.heatmap(pivot, cmap="YlOrRd", linewidths=0.5)
    plt.title("Average Points by Province and Subject (Top 15)")
    plt.tight_layout()
    plt.savefig(r"c:\Users\HP\Desktop\data haraka\heatmap_points.png", dpi=200)
    plt.close()

# Histograms
if "Points" in df.columns:
    plt.figure(figsize=(10, 6))
    df["Points"].dropna().hist(bins=30)
    plt.title("Distribution of Total Points")
    plt.xlabel("Points")
    plt.ylabel("Count")
    plt.tight_layout()
    plt.savefig(r"c:\Users\HP\Desktop\data haraka\hist_points.png", dpi=200)
    plt.close()

if "Choice_Rank" in df.columns:
    plt.figure(figsize=(10, 6))
    df["Choice_Rank"].dropna().hist(bins=30)
    plt.title("Distribution of Choice Rank")
    plt.xlabel("Choice Rank")
    plt.ylabel("Count")
    plt.tight_layout()
    plt.savefig(r"c:\Users\HP\Desktop\data haraka\hist_choice_rank.png", dpi=200)
    plt.close()

print("Analysis complete.")
print("Outputs:")
print("- province_stats.csv / province_stats.xlsx")
print("- subject_stats.csv / subject_stats.xlsx")
print("- heatmap_points.png")
print("- hist_points.png")
print("- hist_choice_rank.png")
