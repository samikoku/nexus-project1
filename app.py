import streamlit as st
import pandas as pd
import plotly.express as px

# --- CONFIGURATION ---
st.set_page_config(page_title="South-South Nexus", layout="wide")

# --- MOCK DATA ---
state_data = {
    "Rivers": {"population": 7200000, "current_igr": 8000000000},
    "Bayelsa": {"population": 2400000, "current_igr": 3000000000},
    "Delta":   {"population": 5600000, "current_igr": 6500000000},
    "Akwa Ibom": {"population": 5400000, "current_igr": 6000000000},
    "Cross River": {"population": 3800000, "current_igr": 4000000000},
    "Edo": {"population": 4700000, "current_igr": 5500000000}
}

# --- SIDEBAR ---
st.sidebar.title("Nexus Control Panel")
st.sidebar.info("Adjust parameters to simulate revenue.")

selected_state = st.sidebar.selectbox("Select State", list(state_data.keys()))
penetration = st.sidebar.slider("Market Penetration (%)", 5, 30, 15)
equity_share = st.sidebar.slider("State Equity Share (%)", 20, 50, 40)
arpu = st.sidebar.number_input("Avg Revenue Per User (NGN)", value=2500)

# --- CALCULATIONS ---
pop = state_data[selected_state]['population']
current_igr = state_data[selected_state]['current_igr']
subscribers = int(pop * (penetration / 100))
monthly_revenue = subscribers * arpu
state_share = monthly_revenue * (equity_share / 100)
annual_igr = state_share * 12

# --- MAIN PAGE ---
st.title(f"Telecom Revenue Roadmap: {selected_state} State")
st.markdown("### Simulating the Financial Impact of State-Backed Telecom Services")

st.markdown("---")

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(label="Monthly IGR Potential", value=f"NGN {state_share:,.0f}")
    
with col2:
    st.metric(label="Annual IGR Potential", value=f"NGN {annual_igr:,.0f}")

with col3:
    st.metric(label="Target Subscribers", value=f"{subscribers:,.0f}")

st.markdown("## Revenue Comparison Analysis")

chart_data = pd.DataFrame({
    'Source': ['Current IGR (Est)', 'Telecom Potential'],
    'Amount': [current_igr, monthly_revenue]
})

fig = px.bar(chart_data, x='Source', y='Amount', color='Source', 
             text_auto=True, 
             color_discrete_map={'Current IGR (Est)': 'grey', 'Telecom Potential': 'green'})

st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.markdown("### Strategy Overview")
st.write(f"""
**Executive Summary for {selected_state} State:**

This analysis demonstrates the potential for {selected_state} State to leverage the new NCC Policy Review. 
By establishing a Special Purpose Vehicle (SPV) holding {equity_share}% equity, the State can capture significant value from the local telecom market.

**Key Projections:**
- **Target Market:** {subscribers:,} citizens (approx {penetration}% of population).
- **Revenue Stream:** Projected monthly contribution of **NGN {state_share:,.0f}** to State coffers.
- **Annual Impact:** A projected **NGN {annual_igr:,.0f}** addition to Internal Generated Revenue.
""")
