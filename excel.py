# Advanced Options Analytics Module
# This module can be integrated with the main dashboard for additional features

import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

class AdvancedOptionsAnalytics:
    """Advanced analytics for options trading"""
    
    def __init__(self, options_data):
        self.data = options_data
        self.setup_data()
    
    def setup_data(self):
        """Prepare data for analysis"""
        if not self.data.empty:
            # Clean and prepare strike prices
            self.data['Strike'] = pd.to_numeric(self.data['Strike'], errors='coerce')
            self.data = self.data.dropna(subset=['Strike'])
            self.data = self.data.sort_values('Strike')
    
    def calculate_option_flow_signals(self):
        """Calculate unusual options activity signals"""
        signals = {}
        
        try:
            # Volume vs OI ratio (indicates fresh money flow)
            if all(col in self.data.columns for col in ['CE_Total_Traded_Volume', 'CE_OI', 'PE_Total_Traded_Volume', 'PE_OI']):
                self.data['CE_Volume_OI_Ratio'] = self.data['CE_Total_Traded_Volume'] / (self.data['CE_OI'] + 1)
                self.data['PE_Volume_OI_Ratio'] = self.data['PE_Total_Traded_Volume'] / (self.data['PE_OI'] + 1)
                
                # High volume/OI ratio indicates unusual activity
                ce_unusual = self.data[self.data['CE_Volume_OI_Ratio'] > self.data['CE_Volume_OI_Ratio'].quantile(0.8)]
                pe_unusual = self.data[self.data['PE_Volume_OI_Ratio'] > self.data['PE_Volume_OI_Ratio'].quantile(0.8)]
                
                signals['unusual_call_activity'] = ce_unusual[['Strike', 'CE_Volume_OI_Ratio', 'CE_Total_Traded_Volume', 'CE_OI']]
                signals['unusual_put_activity'] = pe_unusual[['Strike', 'PE_Volume_OI_Ratio', 'PE_Total_Traded_Volume', 'PE_OI']]
            
            # OI Change analysis
            if 'CE_OI_Change' in self.data.columns and 'PE_OI_Change' in self.data.columns:
                # Significant OI changes
                significant_ce_change = self.data[abs(self.data['CE_OI_Change']) > self.data['CE_OI_Change'].std() * 2]
                significant_pe_change = self.data[abs(self.data['PE_OI_Change']) > self.data['PE_OI_Change'].std() * 2]
                
                signals['significant_call_oi_change'] = significant_ce_change[['Strike', 'CE_OI_Change', 'CE_OI']]
                signals['significant_put_oi_change'] = significant_pe_change[['Strike', 'PE_OI_Change', 'PE_OI']]
        
        except Exception as e:
            st.warning(f"Error calculating option flow signals: {str(e)}")
        
        return signals
    
    def calculate_volatility_skew(self):
        """Calculate volatility skew metrics"""
        try:
            if all(col in self.data.columns for col in ['Strike', 'CE_IV(Spot)', 'PE_IV(Spot)']):
                # ATM strike (closest to current price - we'll use median strike as proxy)
                atm_strike = self.data['Strike'].median()
                
                # Skew calculation
                call_iv_data = self.data[['Strike', 'CE_IV(Spot)']].dropna()
                put_iv_data = self.data[['Strike', 'PE_IV(Spot)']].dropna()
                
                # Calculate skew slope (change in IV per strike)
                if len(call_iv_data) > 1:
                    call_skew_slope, _, _, _, _ = stats.linregress(call_iv_data['Strike'], call_iv_data['CE_IV(Spot)'])
                else:
                    call_skew_slope = 0
                
                if len(put_iv_data) > 1:
                    put_skew_slope, _, _, _, _ = stats.linregress(put_iv_data['Strike'], put_iv_data['PE_IV(Spot)'])
                else:
                    put_skew_slope = 0
                
                # Term structure (if multiple expiries available)
                skew_metrics = {
                    'call_skew_slope': call_skew_slope,
                    'put_skew_slope': put_skew_slope,
                    'atm_strike': atm_strike,
                    'call_iv_range': call_iv_data['CE_IV(Spot)'].max() - call_iv_data['CE_IV(Spot)'].min(),
                    'put_iv_range': put_iv_data['PE_IV(Spot)'].max() - put_iv_data['PE_IV(Spot)'].min()
                }
                
                return skew_metrics
        
        except Exception as e:
            st.warning(f"Error calculating volatility skew: {str(e)}")
            return {}
    
    def calculate_gamma_exposure(self, spot_price=None):
        """Calculate gamma exposure levels"""
        try:
            if not spot_price:
                # Use median strike as proxy for spot
                spot_price = self.data['Strike'].median()
            
            gamma_exposure = {}
            
            if all(col in self.data.columns for col in ['Strike', 'CE_Gamma(Spot)', 'PE_Gamma(Spot)', 'CE_OI', 'PE_OI']):
                # Calculate gamma exposure for each strike
                self.data['CE_Gamma_Exposure'] = self.data['CE_Gamma(Spot)'] * self.data['CE_OI'] * 100  # 100 shares per contract
                self.data['PE_Gamma_Exposure'] = self.data['PE_Gamma(Spot)'] * self.data['PE_OI'] * 100 * -1  # Negative for puts
                self.data['Net_Gamma_Exposure'] = self.data['CE_Gamma_Exposure'] + self.data['PE_Gamma_Exposure']
                
                # Find key gamma levels
                gamma_exposure['total_gamma'] = self.data['Net_Gamma_Exposure'].sum()
                gamma_exposure['max_gamma_strike'] = self.data.loc[self.data['Net_Gamma_Exposure'].abs().idxmax(), 'Strike']
                gamma_exposure['gamma_flip_point'] = self.find_gamma_flip_point()
                
                # Positive/Negative gamma zones
                gamma_exposure['positive_gamma_strikes'] = self.data[self.data['Net_Gamma_Exposure'] > 0]['Strike'].tolist()
                gamma_exposure['negative_gamma_strikes'] = self.data[self.data['Net_Gamma_Exposure'] < 0]['Strike'].tolist()
            
            return gamma_exposure
        
        except Exception as e:
            st.warning(f"Error calculating gamma exposure: {str(e)}")
            return {}
    
    def find_gamma_flip_point(self):
        """Find where gamma changes from positive to negative"""
        try:
            if 'Net_Gamma_Exposure' in self.data.columns:
                # Find the strike where gamma flips sign
                sorted_data = self.data.sort_values('Strike')
                
                for i in range(1, len(sorted_data)):
                    current_gamma = sorted_data.iloc[i]['Net_Gamma_Exposure']
                    previous_gamma = sorted_data.iloc[i-1]['Net_Gamma_Exposure']
                    
                    if (current_gamma > 0 and previous_gamma < 0) or (current_gamma < 0 and previous_gamma > 0):
                        return sorted_data.iloc[i]['Strike']
                
            return None
        except:
            return None
    
    def calculate_dealer_positioning(self):
        """Estimate dealer positioning based on options flow"""
        positioning = {}
        
        try:
            if all(col in self.data.columns for col in ['CE_OI', 'PE_OI', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
                # Assume dealers are short gamma (selling options to retail)
                total_call_oi = self.data['CE_OI'].sum()
                total_put_oi = self.data['PE_OI'].sum()
                
                # Simple dealer position estimation
                positioning['estimated_dealer_call_position'] = -total_call_oi * 0.7  # Assume dealers short 70% of calls
                positioning['estimated_dealer_put_position'] = -total_put_oi * 0.7   # Assume dealers short 70% of puts
                
                # Net dealer position
                positioning['net_dealer_position'] = positioning['estimated_dealer_call_position'] + positioning['estimated_dealer_put_position']
                
                # Hedging pressure
                positioning['call_hedging_pressure'] = total_call_oi / (total_call_oi + total_put_oi)
                positioning['put_hedging_pressure'] = total_put_oi / (total_call_oi + total_put_oi)
        
        except Exception as e:
            st.warning(f"Error calculating dealer positioning: {str(e)}")
        
        return positioning
    
    def create_flow_heatmap(self):
        """Create options flow heatmap"""
        try:
            if all(col in self.data.columns for col in ['Strike', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
                fig = go.Figure()
                
                # Create heatmap data
                strikes = self.data['Strike'].values
                call_volumes = self.data['CE_Total_Traded_Volume'].values
                put_volumes = self.data['PE_Total_Traded_Volume'].values
                
                # Normalize volumes for better visualization
                max_volume = max(call_volumes.max(), put_volumes.max())
                call_volumes_norm = call_volumes / max_volume
                put_volumes_norm = put_volumes / max_volume
                
                # Create heatmap
                fig.add_trace(go.Heatmap(
                    z=[call_volumes_norm, put_volumes_norm],
                    x=strikes,
                    y=['Calls', 'Puts'],
                    colorscale='RdYlGn',
                    showscale=True,
                    hoverongaps=False
                ))
                
                fig.update_layout(
                    title='Options Flow Heatmap',
                    xaxis_title='Strike Price',
                    yaxis_title='Option Type',
                    height=400
                )
                
                return fig
        
        except Exception as e:
            st.warning(f"Error creating flow heatmap: {str(e)}")
            return None
    
    def create_gamma_profile_chart(self):
        """Create gamma exposure profile chart"""
        try:
            if 'Net_Gamma_Exposure' in self.data.columns:
                fig = go.Figure()
                
                # Gamma exposure bars
                colors = ['green' if x >= 0 else 'red' for x in self.data['Net_Gamma_Exposure']]
                
                fig.add_trace(go.Bar(
                    x=self.data['Strike'],
                    y=self.data['Net_Gamma_Exposure'],
                    marker_color=colors,
                    name='Net Gamma Exposure',
                    opacity=0.7
                ))
                
                # Add zero line
                fig.add_hline(y=0, line_dash="dash", line_color="black")
                
                fig.update_layout(
                    title='Gamma Exposure Profile',
                    xaxis_title='Strike Price',
                    yaxis_title='Net Gamma Exposure',
                    height=500,
                    showlegend=False
                )
                
                return fig
        
        except Exception as e:
            st.warning(f"Error creating gamma profile chart: {str(e)}")
            return None
    
    def create_skew_analysis_chart(self):
        """Create volatility skew analysis chart"""
        try:
            if all(col in self.data.columns for col in ['Strike', 'CE_IV(Spot)', 'PE_IV(Spot)']):
                fig = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=['Implied Volatility Skew', 'Put-Call IV Spread'],
                    vertical_spacing=0.1
                )
                
                # IV Skew
                fig.add_trace(
                    go.Scatter(x=self.data['Strike'], y=self.data['CE_IV(Spot)'], 
                              mode='lines+markers', name='Call IV', line=dict(color='green')),
                    row=1, col=1
                )
                
                fig.add_trace(
                    go.Scatter(x=self.data['Strike'], y=self.data['PE_IV(Spot)'], 
                              mode='lines+markers', name='Put IV', line=dict(color='red')),
                    row=1, col=1
                )
                
                # Put-Call IV Spread
                iv_spread = self.data['PE_IV(Spot)'] - self.data['CE_IV(Spot)']
                fig.add_trace(
                    go.Scatter(x=self.data['Strike'], y=iv_spread, 
                              mode='lines+markers', name='Put-Call IV Spread', line=dict(color='purple')),
                    row=2, col=1
                )
                
                fig.add_hline(y=0, line_dash="dash", line_color="black", row=2, col=1)
                
                fig.update_layout(height=600, showlegend=True)
                fig.update_xaxes(title_text="Strike Price", row=2, col=1)
                fig.update_yaxes(title_text="Implied Volatility", row=1, col=1)
                fig.update_yaxes(title_text="IV Spread", row=2, col=1)
                
                return fig
        
        except Exception as e:
            st.warning(f"Error creating skew analysis chart: {str(e)}")
            return None

def display_advanced_analytics(options_data):
    """Main function to display advanced analytics"""
    
    st.header("🔬 Advanced Options Analytics")
    
    if options_data.empty:
        st.warning("No data available for advanced analytics")
        return
    
    # Initialize analytics
    analytics = AdvancedOptionsAnalytics(options_data)
    
    # Create tabs for different analytics
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Option Flow", 
        "📈 Gamma Analysis", 
        "🌊 Volatility Skew", 
        "🏪 Dealer Positioning"
    ])
    
    with tab1:
        st.subheader("Unusual Options Activity")
        
        signals = analytics.calculate_option_flow_signals()
        
        if signals:
            col1, col2 = st.columns(2)
            
            with col1:
                if 'unusual_call_activity' in signals and not signals['unusual_call_activity'].empty:
                    st.write("**🟢 Unusual Call Activity:**")
                    st.dataframe(signals['unusual_call_activity'], use_container_width=True)
                
                if 'significant_call_oi_change' in signals and not signals['significant_call_oi_change'].empty:
                    st.write("**📈 Significant Call OI Changes:**")
                    st.dataframe(signals['significant_call_oi_change'], use_container_width=True)
            
            with col2:
                if 'unusual_put_activity' in signals and not signals['unusual_put_activity'].empty:
                    st.write("**🔴 Unusual Put Activity:**")
                    st.dataframe(signals['unusual_put_activity'], use_container_width=True)
                
                if 'significant_put_oi_change' in signals and not signals['significant_put_oi_change'].empty:
                    st.write("**📉 Significant Put OI Changes:**")
                    st.dataframe(signals['significant_put_oi_change'], use_container_width=True)
        
        # Flow heatmap
        flow_heatmap = analytics.create_flow_heatmap()
        if flow_heatmap:
            st.plotly_chart(flow_heatmap, use_container_width=True)
    
    with tab2:
        st.subheader("Gamma Exposure Analysis")
        
        gamma_exposure = analytics.calculate_gamma_exposure()
        
        if gamma_exposure:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if 'total_gamma' in gamma_exposure:
                    st.metric("Total Net Gamma", f"{gamma_exposure['total_gamma']:,.0f}")
            
            with col2:
                if 'max_gamma_strike' in gamma_exposure:
                    st.metric("Max Gamma Strike", f"₹{gamma_exposure['max_gamma_strike']:,.0f}")
            
            with col3:
                if 'gamma_flip_point' in gamma_exposure and gamma_exposure['gamma_flip_point']:
                    st.metric("Gamma Flip Point", f"₹{gamma_exposure['gamma_flip_point']:,.0f}")
        
        # Gamma profile chart
        gamma_chart = analytics.create_gamma_profile_chart()
        if gamma_chart:
            st.plotly_chart(gamma_chart, use_container_width=True)
        
        # Gamma zones
        if gamma_exposure and 'positive_gamma_strikes' in gamma_exposure:
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**🟢 Positive Gamma Zones:**")
                if gamma_exposure['positive_gamma_strikes']:
                    st.write(f"Strikes: {', '.join(map(str, gamma_exposure['positive_gamma_strikes'][:10]))}")
                else:
                    st.write("None identified")
            
            with col2:
                st.write("**🔴 Negative Gamma Zones:**")
                if gamma_exposure['negative_gamma_strikes']:
                    st.write(f"Strikes: {', '.join(map(str, gamma_exposure['negative_gamma_strikes'][:10]))}")
                else:
                    st.write("None identified")
    
    with tab3:
        st.subheader("Volatility Skew Analysis")
        
        skew_metrics = analytics.calculate_volatility_skew()
        
        if skew_metrics:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if 'call_skew_slope' in skew_metrics:
                    st.metric("Call Skew Slope", f"{skew_metrics['call_skew_slope']:.6f}")
            
            with col2:
                if 'put_skew_slope' in skew_metrics:
                    st.metric("Put Skew Slope", f"{skew_metrics['put_skew_slope']:.6f}")
            
            with col3:
                if 'call_iv_range' in skew_metrics:
                    st.metric("Call IV Range", f"{skew_metrics['call_iv_range']:.4f}")
        
        # Skew analysis chart
        skew_chart = analytics.create_skew_analysis_chart()
        if skew_chart:
            st.plotly_chart(skew_chart, use_container_width=True)
        
        # Skew interpretation
        if skew_metrics:
            st.subheader("📊 Skew Interpretation")
            
            call_slope = skew_metrics.get('call_skew_slope', 0)
            put_slope = skew_metrics.get('put_skew_slope', 0)
            
            if call_slope > 0 and put_slope > 0:
                st.info("📈 **Forward Skew**: Higher strikes have higher IV - indicates fear of upside moves")
            elif call_slope < 0 and put_slope < 0:
                st.info("📉 **Reverse Skew**: Lower strikes have higher IV - indicates fear of downside moves")
            else:
                st.info("⚖️ **Mixed Skew**: Complex volatility structure - analyze individual strikes")
    
    with tab4:
        st.subheader("Dealer Positioning Estimates")
        
        dealer_pos = analytics.calculate_dealer_positioning()
        
        if dealer_pos:
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**📊 Estimated Dealer Positions:**")
                if 'estimated_dealer_call_position' in dealer_pos:
                    st.metric("Dealer Call Position", f"{dealer_pos['estimated_dealer_call_position']:,.0f}")
                if 'estimated_dealer_put_position' in dealer_pos:
                    st.metric("Dealer Put Position", f"{dealer_pos['estimated_dealer_put_position']:,.0f}")
            
            with col2:
                st.write("**⚖️ Hedging Pressure:**")
                if 'call_hedging_pressure' in dealer_pos:
                    st.metric("Call Hedging Pressure", f"{dealer_pos['call_hedging_pressure']:.2%}")
                if 'put_hedging_pressure' in dealer_pos:
                    st.metric("Put Hedging Pressure", f"{dealer_pos['put_hedging_pressure']:.2%}")
        
        # Positioning interpretation
        st.subheader("📊 Positioning Insights")
        
        if dealer_pos and 'net_dealer_position' in dealer_pos:
            net_pos = dealer_pos['net_dealer_position']
            
            if net_pos < -50000:
                st.warning("🔴 **Heavy Dealer Short Position**: Dealers likely hedging by selling underlying - bearish pressure")
            elif net_pos > 50000:
                st.success("🟢 **Heavy Dealer Long Position**: Dealers likely hedging by buying underlying - bullish pressure")
            else:
                st.info("⚖️ **Balanced Dealer Position**: Neutral market maker positioning")

# Usage example:
# display_advanced_analytics(options_dataframe)
