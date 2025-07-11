import streamlit as st
import pandas as pd
from pathlib import Path
import shutil
import tempfile
from generator_core import run_generator

st.set_page_config(page_title="Service Offerings Generator", layout="wide")

st.title("🔧 Service Offerings Generator")
st.markdown("---")

col1, col2 = st.columns([1, 1])

with col1:
    st.header("📁 Input Files")
    uploaded_files = st.file_uploader(
        "Upload ALL_Service_Offering Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more Excel files with naming pattern: ALL_Service_Offering_*.xlsx"
    )
    
    if uploaded_files:
        st.success(f"✅ Uploaded {len(uploaded_files)} file(s)")
        for file in uploaded_files:
            st.text(f"• {file.name}")

with col2:
    st.header("⚙️ Configuration")
    
    tab1, tab2, tab3 = st.tabs(["Basic", "Schedule", "Advanced"])
    
    with tab1:
        st.subheader("Basic Settings")
        
        st.info("""
        **Keywords filtering:**
        - First keyword filters the Parent Offering column
        - Subsequent keywords filter the Name (Child Service Offering lvl 1) column (OR logic - any match)
        """)
        
        keywords = st.text_area(
            "Keywords (one per line)",
            value="",
            help="First keyword filters Parent Offering, others filter Child Service Offering name"
        ).strip().split('\n')
        keywords = [k.strip() for k in keywords if k.strip()]
        
        new_apps = st.text_area(
            "Applications (one per line or comma-separated)",
            value="",
            help="Enter application names - can use newlines or commas to separate"
        ).strip().split('\n')
        new_apps = [a.strip() for a in new_apps if a.strip()]
        
        sr_or_im = st.radio("Service Type", ["SR", "IM"], horizontal=True)
        
        delivery_manager = st.text_input("Delivery Manager", value="")
    
    with tab2:
        st.subheader("Schedule Settings")
        
        schedule_type = st.checkbox("Custom schedule per period")
        
        if not schedule_type:
            schedule_simple = st.text_input("Schedule", value="", placeholder="e.g. Mon-Fri 9-17")
            schedule_suffix = schedule_simple if schedule_simple else ""
        else:
            st.info("📅 Enter schedule periods (e.g., Mon-Thu 9-17, Fri 9-16, Sat 8-12)")
            
            # Allow up to 5 schedule periods
            schedule_parts = []
            num_periods = st.number_input("Number of periods", min_value=1, max_value=5, value=2)
            
            for i in range(num_periods):
                st.markdown(f"**Period {i+1}**")
                col1, col2 = st.columns([2, 1])
                with col1:
                    period = st.text_input(
                        f"Days", 
                        value="", 
                        placeholder="e.g. Mon-Thu or Fri or Sat-Sun",
                        key=f"schedule_period_{i}",
                        label_visibility="collapsed"
                    )
                with col2:
                    hours = st.text_input(
                        f"Hours", 
                        value="", 
                        placeholder="e.g. 9-17",
                        key=f"schedule_hours_{i}",
                        label_visibility="collapsed"
                    )
                
                if period and hours:
                    schedule_parts.append(f"{period} {hours}")
            
            # Join all non-empty schedule parts
            schedule_suffix = " ".join(schedule_parts) if schedule_parts else ""
            
            # Show preview
            if schedule_suffix:
                st.success(f"📋 Schedule: **{schedule_suffix}**")
        
        st.markdown("---")
        col_rsp, col_rsl = st.columns(2)
        with col_rsp:
            rsp_duration = st.text_input("RSP Duration", value="")
        with col_rsl:
            rsl_duration = st.text_input("RSL Duration", value="")
    
    with tab3:
        st.subheader("Advanced Settings")
        
        # Special naming options
        st.markdown("### Naming Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            require_corp = st.checkbox("CORP in Child Service Offerings?")
            
        with col2:
            special_it = st.checkbox("IT Department", disabled=require_corp)
            
        with col3:
            special_hr = st.checkbox("HR Department", disabled=require_corp)
        
        if require_corp:
            delivering_tag = st.text_input(
                "Who delivers the service", 
                value="",
                help="E.g. HS PL, DS DE, IT, Finance, etc."
            )
        else:
            delivering_tag = ""
        
        # Global prod option
        global_prod = st.checkbox("Global Prod", value=False)
        
        # Support groups
        st.markdown("### Support Groups")
        support_group = st.text_input("Support Group", value="")
        managed_by_group = st.text_input(
            "Managed by Group", 
            value="",
            help="Optional - if empty, will use Support Group value"
        )
        
        # Aliases
        st.markdown("### Aliases")
        aliases_on = st.checkbox("Enable Aliases", value=False)  # Default to False
        
        if aliases_on:
            aliases_value = st.text_input(
                "Alias Value",
                value="",
                help="Enter the value to use for aliases"
            )
        else:
            aliases_value = ""

st.markdown("---")

# Show naming examples based on selections
with st.expander("📋 Naming Convention Examples"):
    if 'require_corp' in locals() and require_corp:
        st.markdown("**CORP Example:**")
        st.code("[SR HS PL CORP DS DE] Software assistance Outlook Prod Mon-Fri 8-17")
    elif 'special_it' in locals() and special_it:
        st.markdown("**IT Department Example:**")
        st.code("[SR HS PL IT] Software assistance Outlook Prod Mon-Fri 8-17")
    elif 'special_hr' in locals() and special_hr:
        st.markdown("**HR Department Example:**")
        st.code("[SR HS PL HR] Software assistance Outlook Prod Mon-Fri 8-17")
    else:
        st.markdown("**Standard Example:**")
        st.code("[SR HS PL Permissions] Granting permissions to application Outlook Prod Mon-Fri 9-17")
        st.markdown("From parent: `[Parent HS PL Permissions] Granting permissions to application`")

if st.button("🚀 Generate Service Offerings", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file")
    elif not keywords:
        st.error("⚠️ Please enter at least one keyword")
    elif not new_apps:
        st.error("⚠️ Please enter at least one application")
    elif 'schedule_suffix' not in locals() or not schedule_suffix:
        st.error("⚠️ Please configure schedule")
    else:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                src_dir = Path(temp_dir) / "input"
                out_dir = Path(temp_dir) / "output"
                src_dir.mkdir(exist_ok=True)
                
                for uploaded_file in uploaded_files:
                    file_path = src_dir / uploaded_file.name
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                
                with st.spinner("🔄 Generating service offerings..."):
                    result_file = run_generator(
                        keywords=keywords,
                        new_apps=new_apps,
                        schedule_suffix=schedule_suffix,
                        delivery_manager=delivery_manager,
                        global_prod=global_prod,
                        rsp_duration=rsp_duration,
                        rsl_duration=rsl_duration,
                        sr_or_im=sr_or_im,
                        require_corp=require_corp,
                        delivering_tag=delivering_tag,
                        support_group=support_group,
                        managed_by_group=managed_by_group,
                        aliases_on=aliases_on,
                        aliases_value=aliases_value,
                        src_dir=src_dir,
                        out_dir=out_dir,
                        special_it=special_it if 'special_it' in locals() else False,
                        special_hr=special_hr if 'special_hr' in locals() else False
                    )
                
                st.success("✅ Service offerings generated successfully!")
                
                with open(result_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Generated File",
                        data=f.read(),
                        file_name=result_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.info(f"Generated file: {result_file.name}")
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        Service Offerings Generator v3.0 | With Enhanced Naming Logic
    </div>
    """,
    unsafe_allow_html=True
)