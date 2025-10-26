# sales_dashboard_st_charts.py
import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="PepsiCo Campaign Analysis Dashboard - 14 May 2025 to 14 August 2025", layout="wide")

# ---------------- Helpers ----------------
def human_format(num):
    if pd.isna(num):
        return ""
    try:
        num = float(num)
    except:
        return str(num)
    if abs(num) >= 1_000_000:
        return f'{num/1_000_000:.1f}M'
    elif abs(num) >= 1_000:
        return f'{num/1_000:.0f}K'
    else:
        return str(int(num))

def human_currency(num):
    if pd.isna(num):
        return ""
    try:
        num = float(num)
    except:
        return str(num)
    if abs(num) >= 1_000_000:
        return f'R{num/1_000_000:.1f}M'
    elif abs(num) >= 1_000:
        return f'R{num/1_000:.0f}K'
    else:
        return f'R{num:.0f}'

# ---------------- YOY Analysis ----------------
def yoy_analysis_page():
    st.header("YOY Analysis")    

    # Read Excel
    try:
        df = pd.read_excel('YOY Analysis.xlsx', sheet_name=0, header=1)
        # Normalize expected columns
        df.columns = [c.strip() for c in df.columns]
        expected = ['Product Description', 'QTY Sold Prior Year', 'QTY Sold CAMPAIGN PERIOD', 'Increase in sales from Prior Year AVE']
        if set(expected).issubset(set(df.columns)):
            df = df[expected]
        else:
            # Try to map first 4 columns if names differ
            df = df.iloc[:, :4]
            df.columns = expected
        df = df[df['Product Description'].apply(lambda x: isinstance(x, str))]
    except FileNotFoundError:
        st.error("File `YOY Analysis.xlsx` not found in current directory.")
        return
    except Exception as e:
        st.error(f"Error reading YOY file: {e}")
        return

    # Metrics
    increase_count = (df['Increase in sales from Prior Year AVE'] > 0).sum()
    decrease_count = (df['Increase in sales from Prior Year AVE'] < 0).sum()
    avg_increase = df['Increase in sales from Prior Year AVE'].mean()
    top_grower = df.loc[df['Increase in sales from Prior Year AVE'].idxmax()]
    top_decliner = df.loc[df['Increase in sales from Prior Year AVE'].idxmin()]

    total_prior = df['QTY Sold Prior Year'].sum()
    total_campaign = df['QTY Sold CAMPAIGN PERIOD'].sum()
    total_growth = (total_campaign - total_prior) / total_prior if total_prior != 0 else np.nan

    # Chart 1: YOY pct change
    chart_data1 = pd.DataFrame({
        'Product': df['Product Description'],
        'Increase (%)': df['Increase in sales from Prior Year AVE'] * 100
    })
    st.subheader("YOY Sales Change by Product")
    st.bar_chart(chart_data1.set_index('Product'), height=500)
    
    # Chart 2: Prior vs Campaign volumes
    chart_data2 = pd.DataFrame({
        'Product': df['Product Description'],
        'Prior Year': df['QTY Sold Prior Year'],
        'Campaign Period': df['QTY Sold CAMPAIGN PERIOD']
    })
    st.subheader("YOY Sales Comparison by Product")
    st.bar_chart(chart_data2.set_index('Product'), height=500)

    # Key insights display
    st.subheader("Key Findings")
    st.markdown(f"- {increase_count} products increased, {decrease_count} decreased; average change {avg_increase:.2%}.")
    st.markdown(f"- Top grower: **{top_grower['Product Description']}** ({top_grower['Increase in sales from Prior Year AVE']:.2%}).")
    st.markdown(f"- Top decliner: **{top_decliner['Product Description']}** ({top_decliner['Increase in sales from Prior Year AVE']:.2%}).")
    st.markdown(f"- Total prior: {human_format(total_prior)} → Total campaign: {human_format(total_campaign)} ({total_growth:.2%}).")

# ---------------- Prior Periods ----------------
def prior_periods_page():
    st.header("Campaign Prior Periods Analysis")
    
    # Read Excel
    try:
        df = pd.read_excel('Prior Periods.xlsx', header=1)
        df.columns = [c.strip() for c in df.columns]
        # locate expected columns (fuzzy)
        col_prior = next((c for c in df.columns if 'Prior Year' in c), None)
        col_feb_may = next((c for c in df.columns if 'Feb' in c or '14 Feb' in c), None)
        col_campaign = next((c for c in df.columns if 'CAMPAIGN' in c.upper()), None)
        col_increase = next((c for c in df.columns if 'Increase' in c), None)

        if not (col_prior and col_feb_may and col_campaign and col_increase and 'Product Description' in df.columns):
            st.warning("Could not detect all expected columns automatically. Please ensure the Excel file structure matches the expected layout.")
        df = df[df['Product Description'].apply(lambda x: isinstance(x, str))]
        # Compute avg prior months if possible
        if col_prior and col_feb_may:
            df['Avg Prior Months'] = df[[col_prior, col_feb_may]].mean(axis=1)
        else:
            df['Avg Prior Months'] = np.nan
    except FileNotFoundError:
        st.error("File `Prior Periods.xlsx` not found in current directory.")
        return
    except Exception as e:
        st.error(f"Error reading Prior Periods file: {e}")
        return

    # Chart 1
    chart_data1 = pd.DataFrame({
        'Product': df['Product Description'],
        'Campaign Period Sales': df[col_campaign] if col_campaign else 0,
        'Avg of Prior Months': df['Avg Prior Months']
    })
    st.subheader("Campaign Period Sales vs. Average of Prior Months")
    st.bar_chart(chart_data1.set_index('Product'), height=500)

    # Key findings for Chart 1
    st.subheader("Key Findings - Campaign vs Prior Average")
    products_above_avg = df[df[col_campaign] > df['Avg Prior Months']]['Product Description'].tolist() if col_campaign else []
    products_below_avg = df[df[col_campaign] < df['Avg Prior Months']]['Product Description'].tolist() if col_campaign else []
    st.markdown(f"- {len(products_above_avg)} out of {len(df)} products achieved higher sales during the campaign than their prior average.")
    st.markdown(f"- Products above average likely benefited from campaign activities.")
    st.markdown(f"- Focus on replicating successful tactics and investigating underperformers.")

    # Chart 2: increase %
    chart_data2 = pd.DataFrame({
        'Product': df['Product Description'],
        'Sales Increase (%)': df[col_increase] * 100 if col_increase else 0
    })
    st.subheader("Sales Increase During Campaign vs. Avg of Prior Months")
    st.bar_chart(chart_data2.set_index('Product'), height=500)

    # Key findings
    increase_count = (df[col_increase] > 0).sum() if col_increase else 0
    decrease_count = (df[col_increase] < 0).sum() if col_increase else 0
    avg_increase = df[col_increase].mean() if col_increase else np.nan
    top_grower = df.loc[df[col_increase].idxmax()] if col_increase else None
    top_decliner = df.loc[df[col_increase].idxmin()] if col_increase else None

    st.subheader("Key Findings - Sales Increase Analysis")
    st.markdown(f"- {increase_count} products increased vs avg prior months, {decrease_count} decreased.")
    if top_grower is not None and col_increase:
        st.markdown(f"- Top grower vs prior average: **{top_grower['Product Description']}** (+{top_grower[col_increase]:.1%}).")
    st.markdown(f"- {increase_count} products experienced growth, while {decrease_count} declined.")
    st.markdown(f"- Green bars indicate positive campaign impact.")
    st.markdown(f"- Focus on doubling down on growth products.")

# ---------------- Category Analysis ----------------
def category_analysis_page():
    st.header("Campaign Category Share Analysis")

    # Data (hard-coded)
    data = {
        "Product Description": [
            "BOKOMO CORN FLAKES CEREALS ORIGINAL 1KG",
            "BOKOMO TRADITIONAL OATS 1KG",
            "LIQUI FRUIT 2L",
            "SIMBA CHIPS 120G",
            "WEET BIX CEREALS BOX 450G",
            "WEET BIX CEREALS BOX 900G",
            "WELLINGTONS SWEET CHILLI SAUCE 700ML",
            "WELLINGTONS TOMATO SAUCE 700ml",
            "WHITE STAR INSTANT MAIZE PORRIDGE 1KG",
            "WHITE STAR SUPER MAIZE MEAL MAIZE BAG 2.5KG",
            "WHITE STAR M/MEAL 10KG"
        ],
        "PepsiCo Campaign Category Share": [0.80, 0.14, 0.82, 0.74, 0.79, 0.83, 0.62, 0.45, 0.33, 0.09, 0.06],
        "Competitor Category Share": [0.20, 0.86, 0.18, 0.26, 0.21, 0.17, 0.38, 0.55, 0.67, 0.91, 0.94],
        "Pre-Campaign PepsiCo": [0.76, 0.10, 0.85, 0.73, 0.77, 0.80, 0.69, 0.47, 0.30, 0.12, 0.07],
        "% Change": [0.05, 0.04, -0.03, 0.01, 0.02, 0.02, -0.06, -0.02, 0.03, -0.03, -0.01]
    }
    df = pd.DataFrame(data)

    # Chart 1
    chart_data1 = pd.DataFrame({
        'Product': df["Product Description"],
        'PepsiCo': df["PepsiCo Campaign Category Share"] * 100,
        'Competitors': df["Competitor Category Share"] * 100
    })
    st.subheader("Category Share During Campaign Period by Product")
    st.bar_chart(chart_data1.set_index('Product'), height=500)

    # Chart 2
    chart_data2 = pd.DataFrame({
        'Product': df["Product Description"],
        'Pre-Campaign PepsiCo Share (%)': df["Pre-Campaign PepsiCo"] * 100
    })
    st.subheader("Pre-Campaign PepsiCo Share by Product")
    st.bar_chart(chart_data2.set_index('Product'), height=500)

    # Chart 3
    chart_data3 = pd.DataFrame({
        'Product': df["Product Description"],
        'PepsiCo (Campaign)': df["PepsiCo Campaign Category Share"] * 100,
        'PepsiCo (Pre-Campaign)': df["Pre-Campaign PepsiCo"] * 100,
        'Competitor (Campaign)': df["Competitor Category Share"] * 100
    })
    st.subheader("Category Share Comparison by Product (Campaign vs Pre-Campaign)")
    st.bar_chart(chart_data3.set_index('Product'), height=500)

    # Key findings
    top_gain = df.iloc[df["% Change"].idxmax()]
    top_loss = df.iloc[df["% Change"].idxmin()]
    avg_pepsico_share = df["PepsiCo Campaign Category Share"].mean()
    avg_competitor_share = df["Competitor Category Share"].mean()
    num_gain = (df["% Change"] > 0).sum()
    num_loss = (df["% Change"] < 0).sum()

    st.subheader("Key Findings")
    st.markdown(f"- PepsiCo avg share {avg_pepsico_share:.0%}, competitors {avg_competitor_share:.0%}.")
    st.markdown(f"- {num_gain} products increased share, {num_loss} decreased. Largest gain: **{top_gain['Product Description']}**.")

# ---------------- Campaign Units Analysis ----------------
def campaign_units_page():
    st.header("Campaign Units Analysis (Pre, During, Post)")

    data = {
        "Product Description": [
            "BOKOMO CORN FLAKES CEREALS ORIGINAL 1KG",
            "BOKOMO TRADITIONAL OATS 1KG",
            "LIQUI FRUIT 2L",
            "SIMBA CHIPS 120G",
            "WEET BIX CEREALS BOX 450G",
            "WEET BIX CEREALS BOX 900G",
            "WELLINGTONS SWEET CHILLI SAUCE 700ML",
            "WELLINGTONS TOMATO SAUCE 700ml",
            "WHITE STAR INSTANT MAIZE PORRIDGE 1KG",
            "WHITE STAR SUPER MAIZE MEAL MAIZE BAG 2.5KG",
            "WHITE STAR M/MEAL 10KG"
        ],
        "Pre-Campaign Units (6wks)": [41134, 1116, 43481, 139366, 16784, 21149, 1091, 6336, 34407, 1168, 9259],
        "Campaign Units/Week": [47657, 2146, 53466, 153274, 20570, 21155, 1445, 8003, 36302, 1412, 9409],
        "Post-Campaign Units/Week": [37502, 2115, 44185, 113034, 18148, 17783, 849, 4889, 30976, 1084, 7371],
        "% Change (Campaign vs Pre)": [0.16, 0.92, 0.23, 0.10, 0.23, 0.0003, 0.32, 0.26, 0.06, 0.21, 0.02],
        "% Change (Campaign vs Post)": [-0.21, -0.01, -0.17, -0.26, -0.12, -0.16, -0.41, -0.39, -0.15, -0.23, -0.22]
    }
    df = pd.DataFrame(data)

    # Chart 1
    chart_data1 = pd.DataFrame({
        'Product': df["Product Description"],
        'Pre-Campaign': df["Pre-Campaign Units (6wks)"],
        'Campaign': df["Campaign Units/Week"],
        'Post-Campaign': df["Post-Campaign Units/Week"]
    })
    st.subheader("Units Sold per Product: Pre-, During-, and Post-Campaign")
    st.bar_chart(chart_data1.set_index('Product'), height=500)

    # Key findings for Chart 1
    st.subheader("Key Findings - Units Sold Comparison")
    st.markdown(f"- This chart compares units sold per product before, during, and after the campaign.")
    st.markdown(f"- Campaign period generally drove higher weekly sales across the portfolio.")
    st.markdown(f"- Focus on strategies to extend the positive effects of campaigns.")

    # Chart 2
    chart_data2 = pd.DataFrame({
        'Product': df["Product Description"],
        '% Change (Campaign vs Pre)': df["% Change (Campaign vs Pre)"] * 100
    })
    st.subheader("% Change in Units Sold: Campaign vs Pre-Campaign")
    st.bar_chart(chart_data2.set_index('Product'), height=500)

    # Key findings for Chart 2
    st.subheader("Key Findings - Campaign vs Pre-Campaign")
    num_gain_pre = (df["% Change (Campaign vs Pre)"] > 0).sum()
    num_loss_pre = (df["% Change (Campaign vs Pre)"] < 0).sum()
    top_gain_pre = df.iloc[df["% Change (Campaign vs Pre)"].idxmax()]
    st.markdown(f"- {num_gain_pre} products experienced an increase in units sold during the campaign compared to pre-campaign.")
    st.markdown(f"- Products with strong positive change likely benefited from campaign activities.")
    st.markdown(f"- Replicate successful tactics from top-performing products like {top_gain_pre['Product Description']}.")

    # Chart 3
    chart_data3 = pd.DataFrame({
        'Product': df["Product Description"],
        '% Change (Campaign vs Post)': df["% Change (Campaign vs Post)"] * 100
    })
    st.subheader("% Change in Units Sold: Campaign vs Post-Campaign")
    st.bar_chart(chart_data3.set_index('Product'), height=500)

    # Key findings
    num_gain_post = (df["% Change (Campaign vs Post)"] > 0).sum()
    num_loss_post = (df["% Change (Campaign vs Post)"] < 0).sum()
    top_gain_post = df.iloc[df["% Change (Campaign vs Post)"].idxmax()]
    top_loss_post = df.iloc[df["% Change (Campaign vs Post)"].idxmin()]
    avg_change_pre = df["% Change (Campaign vs Pre)"].mean()
    avg_change_post = df["% Change (Campaign vs Post)"].mean()

    st.subheader("Key Findings - Campaign vs Post-Campaign")
    st.markdown(f"- {num_gain_post} products maintained or increased sales post-campaign compared to the campaign period.")
    st.markdown(f"- Most products saw a drop in sales after the campaign.")
    st.markdown(f"- Develop post-campaign plans to sustain gains.")

    # Summary findings
    st.subheader("Summary Findings")
    st.markdown(f"- Campaign period drove higher weekly sales for most products, but these gains were not always sustained post-campaign.")
    st.markdown(f"- Focus on strategies to extend the positive effects of campaigns beyond the campaign period.")

# ---------------- Campaign Sales Amount Analysis ----------------
def campaign_sales_amount_page():
    st.header("Campaign Sales Amount Analysis (Pre, During, Post)")
    data = {
        "Product Description": [
            "BOKOMO CORN FLAKES CEREALS ORIGINAL 1KG",
            "BOKOMO TRADITIONAL OATS 1KG",
            "LIQUI FRUIT 2L",
            "SIMBA CHIPS 120G",
            "WEET BIX CEREALS BOX 450G",
            "WEET BIX CEREALS BOX 900G",
            "WELLINGTONS SWEET CHILLI SAUCE 700ML",
            "WELLINGTONS TOMATO SAUCE 700ml",
            "WHITE STAR INSTANT MAIZE PORRIDGE 1KG",
            "WHITE STAR SUPER MAIZE MEAL MAIZE BAG 2.5KG",
            "WHITE STAR M/MEAL 10KG"
        ],
        "Pre-Campaign Sales (6wks)": [
            2146039.02, 40101.00, 1959100.98, 2458327.73, 486957.01, 1119207.82,
            53829.88, 194247.98, 993367.85, 49663.65, 1215487.70
        ],
        "Campaign Sales/Week": [
            2512772.12, 81147.63, 2403505.63, 2695830.06, 595658.98, 1127400.03,
            72150.90, 258569.12, 1051719.19, 59890.38, 1154819.48
        ],
        "Post-Campaign Sales/Week": [
            1910801.08, 76227.17, 2006715.47, 2047509.76, 511282.47, 959973.53,
            44785.86, 163818.96, 882593.02, 44569.48, 852939.35
        ],
        "% Change (Campaign vs Pre)": [
            0.17, 1.02, 0.23, 0.10, 0.22, 0.01, 0.34, 0.33, 0.06, 0.21, -0.05
        ],
        "% Change (Campaign vs Post)": [
            -0.24, -0.06, -0.17, -0.24, -0.14, -0.15, -0.38, -0.37, -0.16, -0.26, -0.26
        ]
    }
    df = pd.DataFrame(data)

    # Chart 1
    chart_data1 = pd.DataFrame({
        'Product': df["Product Description"],
        'Pre-Campaign': df["Pre-Campaign Sales (6wks)"],
        'Campaign': df["Campaign Sales/Week"],
        'Post-Campaign': df["Post-Campaign Sales/Week"]
    })
    st.subheader("Sales Amount per Product: Pre-, During-, and Post-Campaign")
    st.bar_chart(chart_data1.set_index('Product'), height=500)

    # Key findings for Chart 1
    st.subheader("Key Findings - Sales Amount Comparison")
    st.markdown(f"- This chart compares sales amounts per product before, during, and after the campaign.")
    st.markdown(f"- Campaign period generally drove higher weekly sales amounts across the portfolio.")
    st.markdown(f"- Focus on strategies to extend the positive effects of campaigns.")

    # Chart 2
    chart_data2 = pd.DataFrame({
        'Product': df["Product Description"],
        '% Change (Campaign vs Pre Sales)': df["% Change (Campaign vs Pre)"] * 100
    })
    st.subheader("% Change in Sales Amount: Campaign vs Pre-Campaign")
    st.bar_chart(chart_data2.set_index('Product'), height=500)

    # Key findings for Chart 2
    st.subheader("Key Findings - Campaign vs Pre-Campaign Sales")
    num_gain_pre = (df["% Change (Campaign vs Pre)"] > 0).sum()
    top_gain_pre = df.iloc[df["% Change (Campaign vs Pre)"].idxmax()]
    st.markdown(f"- {num_gain_pre} products experienced an increase in sales amount during the campaign compared to pre-campaign.")
    st.markdown(f"- Products with strong positive change likely benefited from campaign activities.")
    st.markdown(f"- Replicate successful tactics from top-performing products like {top_gain_pre['Product Description']}.")

    # Chart 3
    chart_data3 = pd.DataFrame({
        'Product': df["Product Description"],
        '% Change (Campaign vs Post Sales)': df["% Change (Campaign vs Post)"] * 100
    })
    st.subheader("% Change in Sales Amount: Campaign vs Post-Campaign")
    st.bar_chart(chart_data3.set_index('Product'), height=500)

    # Key findings
    num_gain_post = (df["% Change (Campaign vs Post)"] > 0).sum()
    st.subheader("Key Findings - Campaign vs Post-Campaign Sales")
    st.markdown(f"- {num_gain_post} products maintained or increased sales post-campaign compared to the campaign period.")
    st.markdown(f"- Many products saw declines post-campaign.")
    st.markdown(f"- Develop post-campaign plans to sustain gains.")

    # Summary findings
    st.subheader("Summary Findings")
    st.markdown(f"- Campaign increased sales for many products but not all gains were sustained post-campaign.")
    st.markdown(f"- Focus on strategies to extend the positive effects of campaigns beyond the campaign period.")

# ---------------- Demographics ----------------
def demographics_page():
    st.header("Shopper Demographics")
    # Data input (hard-coded)
    data = {
        "Product": [
            "BOKOMO CORN FLAKES CEREALS ORIGINAL 1KG",
            "BOKOMO TRADITIONAL OATS 1KG",
            "LIQUI FRUIT 2L",
            "SIMBA CHIPS 120G",
            "WEET BIX CEREALS BOX 450G",
            "WEET BIX CEREALS BOX 900G",
            "WELLINGTONS SWEET CHILLI SAUCE 700ML",
            "WELLINGTONS TOMATO SAUCE 700ml",
            "WHITE STAR INSTANT MAIZE PORRIDGE 1KG",
            "WHITE STAR SUPER MAIZE MEAL MAIZE BAG 2.5KG",
            "WHITE STAR M/MEAL 10KG"
        ],
        "Mon": [10,12,8,9,12,11,8,10,12,16,9],
        "Tue": [14,13,12,14,15,15,14,15,15,16,17],
        "Wed": [13,14,12,11,12,12,11,10,13,10,16],
        "Thu": [18,19,15,14,15,16,14,13,17,12,14],
        "Fri": [22,16,22,18,20,20,17,16,21,17,18],
        "Sat": [18,14,18,18,17,17,18,21,18,16,17],
        "Sun": [4,13,12,16,8,9,18,16,5,12,9],
        "Female": [84,79,66,60,76,74,64,68,79,62,57],
        "Male": [16,21,34,40,24,26,36,32,21,38,43],
        "0-18": [4,5,7,6,4,6,8,9,2,6,5],
        "18-24": [6,6,3,5,6,3,3,3,7,6,5],
        "25-34": [30,17,18,23,28,20,16,17,32,22,18],
        "35-44": [32,40,31,30,32,33,33,35,33,26,35],
        "45-54": [17,18,25,23,18,23,24,20,16,25,23],
        "55-64": [6,8,10,9,8,11,10,9,7,9,10],
        "65+": [4,6,6,4,4,5,6,6,2,7,5]
    }
    df = pd.DataFrame(data)

    days = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    day_means = df[days].mean()
    day_means = day_means / day_means.sum() * 100

    genders = ["Female","Male"]
    gender_means = df[genders].mean()
    gender_means = gender_means / gender_means.sum() * 100

    df["0-24"] = df["0-18"] + df["18-24"]
    age_groups = ["0-24","25-34","35-44","45-54","55-64","65+"]
    age_means = df[age_groups].mean()
    age_means = age_means / age_means.sum() * 100

    top_day = day_means.idxmax(); top_day_pct = day_means.max()
    low_day = day_means.idxmin(); low_day_pct = day_means.min()
    female_pct = gender_means['Female']; male_pct = gender_means['Male']
    top_age_group = age_means.idxmax(); top_age_pct = age_means.max()

    # Plots
    st.subheader("Shoppers by Day of Week (Average % Split)")
    day_data = pd.DataFrame({
        'Day': days,
        'Share (%)': day_means.values
    })
    st.bar_chart(day_data.set_index('Day'), height=400)

    st.markdown(f"**Key Findings - Day of Week:**")
    st.markdown(f"- Shopper activity peaked on **{top_day}** ({top_day_pct:.1f}%) and was lowest on **{low_day}** ({low_day_pct:.1f}%).")
    st.markdown(f"- Friday had the highest activity ({top_day_pct:.1f}%).")

    st.subheader("Gender Breakdown (Average % Split)")
    gender_data = pd.DataFrame({
        'Gender': genders,
        'Share (%)': gender_means.values
    })
    st.bar_chart(gender_data.set_index('Gender'), height=400)

    st.markdown(f"**Key Findings - Gender:**")
    st.markdown(f"- Female shoppers represented **{female_pct:.1f}%** of the total, with males at **{male_pct:.1f}%**.")
    st.markdown(f"- Female shoppers made up {female_pct:.1f}% of shoppers.")

    st.subheader("Age Breakdown (Average % Split)")
    age_data = pd.DataFrame({
        'Age Group': age_groups,
        'Share (%)': age_means.values
    })
    st.bar_chart(age_data.set_index('Age Group'), height=400)

    st.markdown(f"**Key Findings - Age:**")
    st.markdown(f"- The largest age group was **{top_age_group}** ({top_age_pct:.1f}%).")

    st.markdown("### Summary Findings")
    st.markdown(f"- The demographic analysis reveals patterns that should inform timing, targeting, and messaging of future campaigns.")
    st.markdown(f"- Focus on peak shopping days like **{top_day}** for campaign launches.")
    st.markdown(f"- Tailor messaging to resonate with the dominant **{top_age_group}** age group.")

# ---------------- Main App ----------------
def main():
    st.title("PepsiCo Campaign Analysis Dashboard - 14 May 2025 to 14 August 2025")
    with st.sidebar:
        st.header("Navigation")
        menu = st.radio("Select Report", options=[
            "YOY Analysis",
            "Prior Periods",
            "Category Analysis",
            "Campaign Units Analysis",
            "Campaign Sales Amount Analysis",
            "Demographics"
        ])

    if menu == "YOY Analysis":
        yoy_analysis_page()
    elif menu == "Prior Periods":
        prior_periods_page()
    elif menu == "Category Analysis":
        category_analysis_page()
    elif menu == "Campaign Units Analysis":
        campaign_units_page()
    elif menu == "Campaign Sales Amount Analysis":
        campaign_sales_amount_page()
    elif menu == "Demographics":
        demographics_page()

if __name__ == "__main__":
    main()