import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import xlsxwriter

INPUT_CSV = Path("Data for Problem 3.csv")
OUTPUT_XLSX = Path("Sales_Analysis_Problem3.xlsx")


def read_sales_data(csv_path: Path):
    rows = []
    with csv_path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_obj = datetime.strptime(row["Date"], "%Y-%m-%d")
            quantity = int(row["Quantity"])
            price = float(row["Price per Unit"])
            total_amount = float(row["Total Amount"])

            rows.append(
                {
                    "Transaction ID": int(row["Transaction ID"]),
                    "Date": date_obj,
                    "Customer ID": row["Customer ID"],
                    "Gender": row["Gender"],
                    "Age": int(row["Age"]),
                    "Product Category": row["Product Category"],
                    "Quantity": quantity,
                    "Price per Unit": price,
                    "Total Amount": total_amount,
                    # Dataset has no product name; create a practical product key.
                    "Product": f"{row['Product Category']} - ${int(price)}",
                }
            )
    return rows


def aggregate(rows):
    monthly_sales = defaultdict(float)
    category_sales = defaultdict(float)
    product_revenue = defaultdict(float)
    price_quantity_pairs = []
    quarter_sales = defaultdict(float)

    for r in rows:
        month_key = r["Date"].strftime("%Y-%m")
        monthly_sales[month_key] += r["Total Amount"]

        category = r["Product Category"]
        category_sales[category] += r["Total Amount"]

        product_revenue[r["Product"]] += r["Total Amount"]

        price_quantity_pairs.append((r["Price per Unit"], r["Quantity"]))

        quarter = (r["Date"].month - 1) // 3 + 1
        quarter_key = f"{r['Date'].year} Q{quarter}"
        quarter_sales[quarter_key] += r["Total Amount"]

    sorted_months = sorted(monthly_sales.items(), key=lambda x: x[0])
    sorted_categories = sorted(category_sales.items(), key=lambda x: x[1], reverse=True)
    top_products = sorted(product_revenue.items(), key=lambda x: x[1], reverse=True)[:10]
    sorted_quarters = sorted(quarter_sales.items(), key=lambda x: x[0])

    prices = [r["Price per Unit"] for r in rows]
    min_price, max_price = int(min(prices)), int(max(prices))
    bin_size = 50
    bins = list(range((min_price // bin_size) * bin_size, ((max_price // bin_size) + 1) * bin_size + bin_size, bin_size))
    histogram = []
    for i in range(len(bins) - 1):
        low = bins[i]
        high = bins[i + 1]
        count = sum(1 for p in prices if (low <= p < high) or (i == len(bins) - 2 and p == high))
        histogram.append((f"${low}-${high}", count))

    return {
        "monthly": sorted_months,
        "categories": sorted_categories,
        "top_products": top_products,
        "quarters": sorted_quarters,
        "histogram": histogram,
        "price_qty": price_quantity_pairs,
    }


def build_workbook(rows, agg, output_path: Path):
    wb = xlsxwriter.Workbook(output_path.as_posix())

    fmt_header = wb.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1, "align": "center"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    fmt_cell = wb.add_format({"border": 1})
    fmt_money = wb.add_format({"num_format": "$#,##0.00", "border": 1})
    fmt_int = wb.add_format({"num_format": "0", "border": 1})

    ws_data = wb.add_worksheet("Raw Data")
    headers = [
        "Transaction ID",
        "Date",
        "Customer ID",
        "Gender",
        "Age",
        "Product Category",
        "Quantity",
        "Price per Unit",
        "Total Amount",
        "Product",
    ]
    for col, h in enumerate(headers):
        ws_data.write(0, col, h, fmt_header)

    for r_idx, r in enumerate(rows, start=1):
        ws_data.write_number(r_idx, 0, r["Transaction ID"], fmt_int)
        ws_data.write_datetime(r_idx, 1, r["Date"], fmt_date)
        ws_data.write_string(r_idx, 2, r["Customer ID"], fmt_cell)
        ws_data.write_string(r_idx, 3, r["Gender"], fmt_cell)
        ws_data.write_number(r_idx, 4, r["Age"], fmt_int)
        ws_data.write_string(r_idx, 5, r["Product Category"], fmt_cell)
        ws_data.write_number(r_idx, 6, r["Quantity"], fmt_int)
        ws_data.write_number(r_idx, 7, r["Price per Unit"], fmt_money)
        ws_data.write_number(r_idx, 8, r["Total Amount"], fmt_money)
        ws_data.write_string(r_idx, 9, r["Product"], fmt_cell)

    ws_data.set_column("A:A", 14)
    ws_data.set_column("B:B", 12)
    ws_data.set_column("C:C", 12)
    ws_data.set_column("D:D", 10)
    ws_data.set_column("E:E", 8)
    ws_data.set_column("F:F", 18)
    ws_data.set_column("G:G", 10)
    ws_data.set_column("H:I", 14)
    ws_data.set_column("J:J", 22)

    ws_sum = wb.add_worksheet("Summary")

    ws_sum.write("A1", "Monthly Sales", fmt_header)
    ws_sum.write("A2", "Month", fmt_header)
    ws_sum.write("B2", "Revenue", fmt_header)
    for i, (month, sales) in enumerate(agg["monthly"], start=3):
        ws_sum.write(f"A{i}", month, fmt_cell)
        ws_sum.write_number(f"B{i}", sales, fmt_money)

    ws_sum.write("D1", "Category Sales", fmt_header)
    ws_sum.write("D2", "Category", fmt_header)
    ws_sum.write("E2", "Revenue", fmt_header)
    for i, (cat, sales) in enumerate(agg["categories"], start=3):
        ws_sum.write(f"D{i}", cat, fmt_cell)
        ws_sum.write_number(f"E{i}", sales, fmt_money)

    ws_sum.write("G1", "Top Products (Revenue)", fmt_header)
    ws_sum.write("G2", "Product", fmt_header)
    ws_sum.write("H2", "Revenue", fmt_header)
    for i, (product, sales) in enumerate(agg["top_products"], start=3):
        ws_sum.write(f"G{i}", product, fmt_cell)
        ws_sum.write_number(f"H{i}", sales, fmt_money)

    ws_sum.write("J1", "Price Distribution", fmt_header)
    ws_sum.write("J2", "Price Range", fmt_header)
    ws_sum.write("K2", "Count", fmt_header)
    for i, (price_range, count) in enumerate(agg["histogram"], start=3):
        ws_sum.write(f"J{i}", price_range, fmt_cell)
        ws_sum.write_number(f"K{i}", count, fmt_int)

    ws_sum.write("M1", "Quarterly Sales", fmt_header)
    ws_sum.write("M2", "Quarter", fmt_header)
    ws_sum.write("N2", "Revenue", fmt_header)
    for i, (q, sales) in enumerate(agg["quarters"], start=3):
        ws_sum.write(f"M{i}", q, fmt_cell)
        ws_sum.write_number(f"N{i}", sales, fmt_money)

    ws_scatter = wb.add_worksheet("Price_vs_Quantity")
    ws_scatter.write("A1", "Price per Unit", fmt_header)
    ws_scatter.write("B1", "Quantity", fmt_header)
    for i, (price, qty) in enumerate(agg["price_qty"], start=2):
        ws_scatter.write_number(f"A{i}", price)
        ws_scatter.write_number(f"B{i}", qty)

    ws_charts = wb.add_worksheet("Charts")
    ws_charts.hide_gridlines(2)

    month_start = 3
    month_end = month_start + len(agg["monthly"]) - 1
    max_month_idx = max(range(len(agg["monthly"])), key=lambda i: agg["monthly"][i][1])
    min_month_idx = min(range(len(agg["monthly"])), key=lambda i: agg["monthly"][i][1])
    points = []
    for i in range(len(agg["monthly"])):
        if i == max_month_idx:
            points.append({"fill": {"color": "#16A34A"}, "border": {"color": "#166534"}})
        elif i == min_month_idx:
            points.append({"fill": {"color": "#DC2626"}, "border": {"color": "#991B1B"}})
        else:
            points.append({"fill": {"color": "#3B82F6"}, "border": {"color": "#1D4ED8"}})

    chart_monthly = wb.add_chart({"type": "line"})
    chart_monthly.add_series(
        {
            "name": "Monthly Revenue",
            "categories": f"=Summary!$A${month_start}:$A${month_end}",
            "values": f"=Summary!$B${month_start}:$B${month_end}",
            "marker": {"type": "circle", "size": 7},
            "points": points,
            "line": {"color": "#2563EB", "width": 2.25},
        }
    )
    chart_monthly.set_title({"name": "1) Monthly Sales Trend (Peak/Dip Highlighted)"})
    chart_monthly.set_y_axis({"name": "Revenue"})
    chart_monthly.set_legend({"none": True})

    cat_start = 3
    cat_end = cat_start + len(agg["categories"]) - 1
    chart_cat = wb.add_chart({"type": "column"})
    chart_cat.add_series(
        {
            "name": "Category Revenue",
            "categories": f"=Summary!$D${cat_start}:$D${cat_end}",
            "values": f"=Summary!$E${cat_start}:$E${cat_end}",
            "fill": {"color": "#0EA5E9"},
            "border": {"color": "#0369A1"},
            "data_labels": {"value": True},
        }
    )
    chart_cat.set_title({"name": "2) Category-wise Sales Comparison"})
    chart_cat.set_y_axis({"name": "Revenue"})
    chart_cat.set_legend({"none": True})

    prod_start = 3
    prod_end = prod_start + len(agg["top_products"]) - 1
    chart_products = wb.add_chart({"type": "bar"})
    chart_products.add_series(
        {
            "name": "Top Products by Revenue",
            "categories": f"=Summary!$G${prod_start}:$G${prod_end}",
            "values": f"=Summary!$H${prod_start}:$H${prod_end}",
            "fill": {"color": "#14B8A6"},
            "border": {"color": "#0F766E"},
            "data_labels": {"value": True},
        }
    )
    chart_products.set_title({"name": "3) Top Selling Products (Revenue)"})
    chart_products.set_x_axis({"name": "Revenue"})
    chart_products.set_legend({"none": True})

    hist_start = 3
    hist_end = hist_start + len(agg["histogram"]) - 1
    chart_hist = wb.add_chart({"type": "column"})
    chart_hist.add_series(
        {
            "name": "Count",
            "categories": f"=Summary!$J${hist_start}:$J${hist_end}",
            "values": f"=Summary!$K${hist_start}:$K${hist_end}",
            "fill": {"color": "#F59E0B"},
            "border": {"color": "#B45309"},
        }
    )
    chart_hist.set_title({"name": "4) Price Distribution (Histogram)"})
    chart_hist.set_x_axis({"name": "Price Range"})
    chart_hist.set_y_axis({"name": "Transactions"})
    chart_hist.set_legend({"none": True})

    q_start = 3
    q_end = q_start + len(agg["quarters"]) - 1
    chart_season = wb.add_chart({"type": "area"})
    chart_season.add_series(
        {
            "name": "Quarterly Revenue",
            "categories": f"=Summary!$M${q_start}:$M${q_end}",
            "values": f"=Summary!$N${q_start}:$N${q_end}",
            "fill": {"color": "#A78BFA", "transparency": 20},
            "border": {"color": "#6D28D9"},
        }
    )
    chart_season.set_title({"name": "5) Seasonal Sales Variation"})
    chart_season.set_y_axis({"name": "Revenue"})
    chart_season.set_legend({"none": True})

    chart_pie = wb.add_chart({"type": "pie"})
    chart_pie.add_series(
        {
            "name": "Revenue Share by Category",
            "categories": f"=Summary!$D${cat_start}:$D${cat_end}",
            "values": f"=Summary!$E${cat_start}:$E${cat_end}",
            "data_labels": {"percentage": True, "category": True},
        }
    )
    chart_pie.set_title({"name": "6) Revenue Contribution by Category"})

    scatter_end = 1 + len(agg["price_qty"])
    chart_scatter = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
    chart_scatter.add_series(
        {
            "name": "Price vs Quantity",
            "categories": f"=Price_vs_Quantity!$A$2:$A${scatter_end}",
            "values": f"=Price_vs_Quantity!$B$2:$B${scatter_end}",
            "marker": {"type": "circle", "size": 5, "fill": {"color": "#F43F5E"}, "border": {"color": "#9F1239"}},
            "line": {"none": True},
        }
    )
    chart_scatter.set_title({"name": "7) Price vs Quantity Scatter Plot"})
    chart_scatter.set_x_axis({"name": "Price per Unit"})
    chart_scatter.set_y_axis({"name": "Quantity Sold"})
    chart_scatter.set_legend({"none": True})

    ws_charts.insert_chart("A1", chart_monthly, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("I1", chart_cat, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("A20", chart_products, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("I20", chart_hist, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("A39", chart_season, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("I39", chart_pie, {"x_scale": 1.12, "y_scale": 1.15})
    ws_charts.insert_chart("A58", chart_scatter, {"x_scale": 1.12, "y_scale": 1.15})

    wb.close()


def main():
    rows = read_sales_data(INPUT_CSV)
    agg = aggregate(rows)
    build_workbook(rows, agg, OUTPUT_XLSX)
    print(f"Created: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
