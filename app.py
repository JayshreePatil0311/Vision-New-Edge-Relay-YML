import base64
import io
import os
import pandas as pd
import streamlit as st
import yaml

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Anvex.ai - Vision Camera Automation Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# -----------------------------------------------------------------------------
# CUSTOM DARK MODE CSS
# -----------------------------------------------------------------------------
st.markdown(
    """
<style>
    /* Force Dark Theme Background */
    .stApp {
        background-color: #080B10;
        color: #E2E8F0;
    }
    
    /* Header Card Styling */
    .header-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
    }
    
    .header-title-container {
        display: flex;
        align-items: center;
        gap: 16px;
    }

    .header-logo-img {
        height: 48px;
        width: auto;
        object-fit: contain;
    }
    
    .header-title {
        color: #38BDF8;
        font-size: 28px;
        font-weight: 700;
        margin: 0;
    }
    
    .header-subtitle {
        color: #94A3B8;
        font-size: 14px;
        margin-top: 4px;
    }
    
    /* Feature Cards */
    .feature-card {
        background-color: #0F172A;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .feature-title {
        color: #38BDF8;
        font-weight: 600;
        font-size: 15px;
        margin-bottom: 6px;
    }
    .feature-desc {
        color: #64748B;
        font-size: 12px;
    }
    
    /* Mandatory Red Star Label */
    .required-star {
        color: #EF4444;
        font-weight: bold;
    }
    
    /* Input & Table Styling */
    .stDataFrame {
        border: 1px solid #334155;
        border-radius: 8px;
    }
    
    /* Custom Button Styling */
    .stButton>button {
        background: linear-gradient(90deg, #0EA5E9 0%, #0284C7 100%);
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 600;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #0284C7 0%, #0369A1 100%);
    }

    /* SOP Section Styling */
    .sop-header {
        color: #38BDF8;
        border-bottom: 2px solid #0EA5E9;
        padding-bottom: 6px;
        margin-top: 24px;
        margin-bottom: 12px;
    }
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# HELPER: LOAD LOCAL LOGO AS BASE64
# -----------------------------------------------------------------------------
def get_image_base64(image_path):
  if os.path.exists(image_path):
    with open(image_path, "rb") as img_file:
      ext = image_path.split(".")[-1].lower()
      mime = "image/png" if ext == "png" else "image/jpeg"
      encoded = base64.b64encode(img_file.read()).decode()
      return f"data:{mime};base64,{encoded}"
  return None


# Check for logo.png or logo.jpg in local directory
logo_src = get_image_base64("logo.png") or get_image_base64("logo.jpg")

if logo_src:
  logo_html = (
      f'<img src="{logo_src}" class="header-logo-img" alt="Company Logo"/>'
  )
else:
  logo_html = '<span style="font-size:28px;">⚡</span>'

# -----------------------------------------------------------------------------
# HEADER UI (No Sidebar)
# -----------------------------------------------------------------------------
st.markdown(
    f"""
<div class="header-card">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div class="header-title-container">
            {logo_html}
            <div>
                <div class="header-title">Anvex.ai</div>
                <div class="header-subtitle">Vision Camera Automation Suite</div>
            </div>
        </div>
        <div style="background: #0284C722; border: 1px solid #0284C7; color: #38BDF8; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 600;">
            👁 AI Vision Platform
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Feature Badges
col1, col2, col3, col4 = st.columns(4)
with col1:
  st.markdown(
      '<div class="feature-card"><div class="feature-title">📄 Excel'
      ' Generator</div><div class="feature-desc">Builds standard camera config'
      " sheets automatically</div></div>",
      unsafe_allow_html=True,
  )
with col2:
  st.markdown(
      '<div class="feature-card"><div class="feature-title">⚡ streams.yml'
      ' Converter</div><div class="feature-desc">Converts uploaded Excel into'
      " deployment YAML format</div></div>",
      unsafe_allow_html=True,
  )
with col3:
  st.markdown(
      '<div class="feature-card"><div class="feature-title">🚀 Edge Relay'
      ' Ready</div><div class="feature-desc">Essential for Edge Relay server'
      " deployments</div></div>",
      unsafe_allow_html=True,
  )
with col4:
  st.markdown(
      '<div class="feature-card"><div class="feature-title">🔒 RTSP / RTMP'
      ' Validated</div><div class="feature-desc">Sanitizes URLs, ports, &'
      " credentials automatically</div></div>",
      unsafe_allow_html=True,
  )

st.write("")

# -----------------------------------------------------------------------------
# TABS: MODULE 1, MODULE 2 & MODULE 3
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs([
    "📄 Module 1: Camera Setup & Excel Generator",
    "⚙️ Module 2: Upload Excel & Generate YAML",
    "📚 Module 3: Deployment SOP & Troubleshooting",
])

# --- MODULE 1: DYNAMIC EXCEL TEMPLATE GENERATOR ---
with tab1:
  st.subheader("Generate Camera Configuration Excel Template")
  st.write(
      "Specify the number of cameras and Client ID to generate a pre-formatted"
      " Excel template."
  )

  # Inputs: Number of Cameras & Client ID side-by-side with Red Stars
  input_col1, input_col2 = st.columns(2)

  with input_col1:
    st.markdown(
        'Number of Cameras <span class="required-star">*</span>',
        unsafe_allow_html=True,
    )
    num_cameras = st.number_input(
        "Number of Cameras",
        min_value=1,
        max_value=500,
        value=None,
        placeholder="Enter camera count (e.g. 5)",
        label_visibility="collapsed",
    )

  with input_col2:
    st.markdown(
        'Client ID <span class="required-star">*</span>', unsafe_allow_html=True
    )
    client_id = st.text_input(
        "Client ID",
        value="",
        placeholder="e.g. 6a8d6514a73283244b121f79",
        label_visibility="collapsed",
    )

  # Validation check
  if num_cameras is None or not client_id.strip():
    st.info(
        "💡 Please fill in both **Number of Cameras** and **Client ID** above"
        " to generate the Excel template."
    )
  else:
    # Generate Dynamic Data
    cam_ids = [f"cam_{i+1}" for i in range(int(num_cameras))]
    rtsp_urls = [""] * int(num_cameras)  # Left blank for manual entry
    rtmp_urls = [
        f"rtmp://vision-media-server.anvex.ai:1935/{client_id.strip()}/{cam_id}"
        for cam_id in cam_ids
    ]

    template_df = pd.DataFrame({
        "Camera Number": cam_ids,
        "RTSP Link": rtsp_urls,
        "RTMP URL": rtmp_urls,
    })

    st.write("")
    st.info(
        "📌 **Note**: Please fill the **RTSP Link** column manually according"
        " to your camera brand URL format (e.g., Hikvision, Dahua, CP Plus,"
        " etc.)."
    )

    # Optional Preview Checkbox
    show_preview = st.checkbox("🔍 Show Excel Template Preview")
    if show_preview:
      st.markdown("### 🔍 Step 1: Excel Template Preview")
      st.dataframe(template_df, use_container_width=True)

    # Prepare Download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
      template_df.to_excel(writer, index=False, sheet_name="Streams")

    st.markdown("### 💾 Download Template")
    st.download_button(
        label="📥 Download Excel Template",
        data=buffer.getvalue(),
        file_name=f"camera_config_{client_id.strip()[:8]}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

# --- MODULE 2: UPLOAD & YAML CONVERTER ---
with tab2:
  st.subheader("Upload Excel Sheet, Validate & Build `streams.yml`")
  st.write(
      "Upload the completed Excel configuration sheet. The system will verify"
      " all camera records before generating the YAML configuration."
  )

  uploaded_file = st.file_uploader(
      "Upload Excel Configuration Sheet", type=["xlsx", "xls"]
  )

  if uploaded_file:
    try:
      df = pd.read_excel(uploaded_file)

      # Clean column names
      df.columns = [str(col).strip() for col in df.columns]

      # Map columns flexibly
      cam_col = next(
          (c for c in df.columns if "camera" in c.lower() or "id" in c.lower()),
          df.columns[0],
      )
      rtsp_col = next(
          (c for c in df.columns if "rtsp" in c.lower() or "link" in c.lower()),
          df.columns[1],
      )
      rtmp_col = next(
          (
              c
              for c in df.columns
              if "rtmp" in c.lower() or "target" in c.lower()
          ),
          df.columns[2],
      )

      st.success(
          "✅ Excel File Loaded Successfully! Found"
          f" **{len(df)} camera streams**."
      )

      # 1. Excel Preview Section
      st.markdown("### 🔍 Step 1: Validate Excel Data")
      st.dataframe(
          df[[cam_col, rtsp_col, rtmp_col]].rename(
              columns={
                  cam_col: "Camera ID",
                  rtsp_col: "RTSP Source URL",
                  rtmp_col: "RTMP Target URL",
              }
          ),
          use_container_width=True,
      )

      # 2. Build YAML Object
      streams_list = []
      for _, row in df.iterrows():
        cam_id = (
            str(row[cam_col]).strip() if pd.notna(row[cam_col]) else ""
        )
        rtsp_url = (
            str(row[rtsp_col]).strip() if pd.notna(row[rtsp_col]) else ""
        )
        rtmp_url = (
            str(row[rtmp_col]).strip() if pd.notna(row[rtmp_col]) else ""
        )

        stream_entry = {
            "id": cam_id,
            "source": rtsp_url,
            "transport": "rtmp",
            "target": rtmp_url,
            "fallback": rtmp_url,
            "srt_streamid": cam_id,
        }
        streams_list.append(stream_entry)

      yaml_data = {
          "server": {"listen": "127.0.0.1:8080", "log_level": "info"},
          "defaults": {
              "rtsp_transport": "tcp",
              "srt_latency_ms": 200,
              "srt_connect_timeout_ms": 3000,
              "srt_payload_size": 1316,
              "on_stall_restart_s": 15,
              "startup_timeout_s": 10,
              "restart_backoff_base_s": 1,
              "restart_backoff_max_s": 60,
              "healthy_reset_s": 60,
              "srt_failure_threshold": 3,
              "rtmp_hold_s": 60,
              "rtmp_hold_max_s": 600,
              "kill_grace_s": 5,
          },
          "streams": streams_list,
      }

      # Build YAML Header Comment
      yaml_header = """# edge-relay stream configuration — the single source of truth.
#
# Validated on load and on every reload (SIGHUP, or simply saving this file
# — changes are picked up automatically within ~1s via inotify). An invalid
# file is rejected with an error in the logs; already-running streams are
# left untouched.
#
# Copy this file to streams.yaml (install.sh does this automatically on
# first install) and edit it for your cameras.

"""
      yaml_body = yaml.dump(
          yaml_data, sort_keys=False, default_flow_style=False
      )
      full_yaml_str = yaml_header + yaml_body

      # 3. YAML Code Preview Section
      st.markdown("### 📝 Step 2: Review Generated YAML Code")
      st.code(full_yaml_str, language="yaml")

      # 4. Download Option
      st.markdown("### 💾 Step 3: Download Configuration")
      st.download_button(
          label="📥 Download streams.yml",
          data=full_yaml_str,
          file_name="streams.yml",
          mime="text/yaml",
      )

    except Exception as e:
      st.error(f"Error processing the Excel file: {str(e)}")

# --- MODULE 3: DEPLOYMENT SOP & TROUBLESHOOTING ---
with tab3:
  st.title("ANVEX AI VISION")
  st.subheader("EDGE RELAY – CAMERA CONFIGURATION & DEPLOYMENT SOP")
  st.info(
      "**Workflow Overview:** RTSP Camera/NVR/DVR Onboarding → VLC Validation"
      " → streams.yaml → Edge Relay → RTMP Monitoring"
  )

  # 1. Purpose
  st.markdown("### 1. Purpose")
  st.write(
      "This SOP defines the complete process for onboarding client cameras"
      " into Anvex Vision Edge Relay. It covers the information to collect"
      " from the client, RTSP URL construction and testing, camera/NVR/DVR"
      " encoding precautions, brand-wise main/sub-stream selection, RTMP URL"
      " generation, streams.yaml creation, validation, and troubleshooting."
  )

  # 2. End-to-End Workflow
  st.markdown("### 2. End-to-End Workflow")
  workflow_data = {
      "Step": list(range(1, 11)),
      "Activity": [
          "Collect client camera/NVR/DVR information and credentials",
          "Confirm camera/NVR/DVR network reachability",
          "Construct RTSP URL",
          "Test RTSP in VLC (Ctrl+N)",
          "Fix encoding / RTSP settings if required",
          "Assign Anvex Camera ID and Client ID",
          "Generate RTMP target URL",
          "Prepare and validate streams.yaml",
          "Deploy Edge Relay",
          "Verify dashboard and RTMP output",
      ],
      "Output": [
          "Camera inventory",
          "Ping/port result",
          "RTSP URL",
          "Working / Not Working",
          "Compatible stream",
          "cam01, cam02, etc.",
          "RTMP URL",
          "Validated YAML",
          "Running relay",
          "Deployment sign-off",
      ],
  }
  st.table(pd.DataFrame(workflow_data))

  # 3. Information Required From Client Before Configuration
  st.markdown("### 3. Information Required From Client Before Configuration")
  st.write(
      "The following information must be collected for every agreed camera. Do"
      " not start YAML generation until the required fields are available."
  )
  info_data = {
      "Required Item": [
          "Client Name",
          "Client ID",
          "Camera ID",
          "Camera Name / Location",
          "Device Type",
          "Brand",
          "Camera/NVR/DVR IP",
          "RTSP Port",
          "Camera Username",
          "Camera Password",
          "Channel No.",
          "Working RTSP URL",
          "NVR/DVR Web Credentials",
          "Site Contact",
      ],
      "What to collect": [
          "Official client/site name",
          "Anvex-provided unique Client ID",
          "Anvex ID such as cam01, cam02, or agreed naming",
          "Reception, Gate, Production Floor, etc.",
          "IP Camera / NVR / DVR",
          "CP Plus / Hikvision / Dahua / UNV / Axis / Other",
          "Local/private IP address",
          "Usually 554, but verify actual value",
          "RTSP-capable device account",
          "RTSP-capable device account password",
          "Required when stream is exposed through NVR/DVR",
          "Final tested URL",
          "Only if Anvex team is authorized to configure the recorder",
          "Client technical contact",
      ],
      "Mandatory?": [
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes",
          "Yes for NVR/DVR",
          "Yes",
          "If required",
          "Recommended",
      ],
  }
  st.table(pd.DataFrame(info_data))

  # 4. Example: Client Excel Required Structure
  st.markdown("### 4. For Example : Client Excel – Required Structure")
  excel_struct_data = {
      "Column": [
          "Client ID",
          "Camera ID",
          "Camera Name",
          "Device Type",
          "Brand",
          "Model",
          "Camera Channel",
          "Camera IP / NVR IP",
          "RTSP Port",
          "Username",
          "Password",
          "Main RTSP URL",
          "Sub RTSP URL",
          "RTSP Test",
          "Main Stream Settings",
          "Sub Stream Settings",
          "RTMP URL",
      ],
      "Example": [
          "6a8d6514a73283244b121f79",
          "cam01",
          "Reception Camera",
          "NVR",
          "Hikvision",
          "DS-xxxx",
          "1",
          "192.168.1.101",
          "554",
          "admin",
          "<secure value>",
          "rtsp://.../101",
          "rtsp://.../102",
          "PASS",
          "1080p / 15 FPS / H.264 / CBR / 2500 Kbps",
          "720p / 10–15 FPS / H.264 / lower bitrate",
          (
              "rtmp://vision-media-server.anvex.ai:1935/<ClientID>/<CameraID>"
          ),
      ],
  }
  st.table(pd.DataFrame(excel_struct_data))

  # 5. Network & Camera Pre-Deployment Precautions
  st.markdown("### 5. Network & Camera Pre-Deployment Precautions")
  st.markdown("""
    - Prefer wired Ethernet between the camera/NVR and the deployment network. Avoid Wi-Fi where possible.
    - Use TCP transport for RTSP in Edge Relay where supported: `rtsp_transport: tcp`.
    - Use static/reserved IP addresses for NVRs and IP cameras so the source address does not unexpectedly change.
    - Plan approximately 2.5–3.0 Mbps upload capacity per camera as a baseline for a 1080p/H.264 stream at the recommended settings; actual bandwidth depends on scene complexity and encoder settings.
    - Confirm firewall rules permit the local RTSP connection and outbound connection to the Anvex media server.
    """)

  # 6. Recommended Video Encoding Settings
  st.markdown("### 6. Recommended Video Encoding Settings")
  enc_data = {
      "Parameter": [
          "Resolution",
          "FPS",
          "GOP / Keyframe",
          "Bitrate",
          "Bitrate Type",
          "Codec",
          "Audio",
          "Profile",
      ],
      "Main Stream – Recommended": [
          "1920×1080 (1080p)",
          "15 FPS",
          "15",
          "≈2500 Kbps",
          "CBR",
          "H.264",
          "Disable if not required",
          "Baseline/Main where available",
      ],
      "Sub Stream – Recommended": [
          "1280×720 or lower",
          "10–15 FPS",
          "15",
          "≈500–1200 Kbps",
          "CBR",
          "H.264",
          "Disable if not required",
          "Baseline/Main where available",
      ],
      "Why": [
          "Main stream retains detail; sub stream reduces bandwidth/processing.",
          "Stable processing and predictable bandwidth.",
          "Frequent keyframes improve stream recovery and seeking.",
          "Keeps network load controlled; tune by scene/camera.",
          "More predictable bandwidth than VBR.",
          "Broad compatibility with FFmpeg/VLC/media pipelines.",
          "Reduces unnecessary bandwidth/compatibility issues.",
          "Prefer broadly compatible H.264 profiles.",
      ],
  }
  st.table(pd.DataFrame(enc_data))
  st.caption(
      "**Important**: These are Edge Relay compatibility recommendations, not"
      " universal camera requirements. Exact limits differ by camera model."
      " Verify the model's supported resolution/FPS/codec ranges before"
      " changing settings."
  )

  # 7. Main Stream vs Sub Stream – What Changes?
  st.markdown("### 7. Main Stream vs Sub Stream – What Changes?")
  st.write(
      "Use the main stream when the Vision use case needs higher visual detail."
      " Use the sub stream when lower bandwidth or lower decode load is"
      " preferred. The stream number/path is vendor-specific."
  )
  stream_change_data = {
      "Vendor / Family": [
          "Hikvision NVR/DVR/IPC",
          "Hikvision channel 2",
          "Dahua / compatible CP Plus",
          "UNV",
      ],
      "Main Stream": [
          "…/Streaming/channels/101",
          "…/channels/201",
          "…channel=1&subtype=0",
          "…/unicast/c1/s0/live",
      ],
      "Sub Stream": [
          "…/Streaming/channels/102",
          "…/channels/202",
          "…channel=1&subtype=1",
          "…/unicast/c1/s1/live",
      ],
      "Meaning": [
          "01 = main; 02 = sub for channel 1.",
          "20 = channel 2; final digit identifies stream.",
          "subtype 0 = main; subtype 1 = sub.",
          "s0 = main; s1 = sub.",
      ],
  }
  st.table(pd.DataFrame(stream_change_data))
  st.warning(
      "Do not blindly replace 0/1 across brands. In one vendor, 0/1 can"
      " identify stream type; in another, 01/02 can be part of the"
      " channel/stream path. Always use the exact path supported by the"
      " device."
  )

  # 8. Brand-Wise RTSP URL Formats
  st.markdown("### 8. Brand-Wise RTSP URL Formats")
  brand_data = {
      "Brand": ["Hikvision", "Dahua", "CP Plus", "UNV", "Axis"],
      "Main Stream Example": [
          "rtsp://USER:PASS@IP:554/Streaming/channels/101",
          "rtsp://USER:PASS@IP:554/cam/realmonitor?channel=1&subtype=0",
          (
              "Use model/NVR-supported RTSP path; commonly Dahua-compatible"
              " patterns exist"
          ),
          "rtsp://USER:PASS@IP:554/unicast/c1/s0/live",
          "rtsp://USER:PASS@IP:554/axis-media/media.amp",
      ],
      "Sub Stream Example": [
          "rtsp://USER:PASS@IP:554/Streaming/channels/102",
          "rtsp://USER:PASS@IP:554/cam/realmonitor?channel=1&subtype=1",
          "Use model/NVR-supported sub-stream path",
          "rtsp://USER:PASS@IP:554/unicast/c1/s1/live",
          "Model/config dependent",
      ],
      "Notes": [
          "Channel 1: 101 main, 102 sub.",
          "Channel starts at 1; subtype 0/1.",
          "Verify exact model/manual before assuming path.",
          "c1 = channel 1; s0/s1 = stream.",
          "Confirm exact stream profile/path.",
      ],
  }
  st.table(pd.DataFrame(brand_data))

  # 9. Password / URL Encoding Rule
  st.markdown("### 9. Password / URL Encoding Rule")
  st.write(
      "If the username or password contains reserved URL characters, encode"
      " them before placing them inside an RTSP URL."
  )
  enc_rule_data = {
      "Character": ["@", "#", ":", "%", "&"],
      "URL-Encoded": ["%40", "%23", "%3A", "%25", "%26"],
      "Example": [
          "P@ss → P%40ss",
          "P#ss → P%23ss",
          "P:ss → P%3Ass",
          "P%ss → P%25ss",
          "P&ss → P%26ss",
      ],
  }
  st.table(pd.DataFrame(enc_rule_data))
  st.caption(
      "If VLC fails to authenticate, check the credentials and URL encoding"
      " before changing camera settings."
  )

  # 10. RTSP Validation – Mandatory Before YAML
  st.markdown("### 10. RTSP Validation – Mandatory Before YAML")
  st.markdown("""
    1. Open **VLC Media Player** on the deployment/support machine.
    2. Go to **Media → Open Network Stream**, or press `Ctrl+N`.
    3. Paste the complete RTSP URL.
    4. Click **Play**.
    5. Confirm that live video appears and remains stable for at least 1–2 minutes.
    6. Record the result in the Excel inventory as **PASS** or **FAIL**.
    7. If the feed fails, troubleshoot credentials, port, path, encoding, network reachability, and vendor-specific settings before creating the final YAML.
    """)

  # 11. If RTSP Does Not Work – Troubleshooting Sequence
  st.markdown("### 11. If RTSP Does Not Work – Troubleshooting Sequence")
  trouble_seq_data = {
      "Check": [
          "1. Device reachable",
          "2. RTSP port",
          "3. Credentials",
          "4. URL path",
          "5. URL encoding",
          "6. Codec",
          "7. FPS/GOP",
          "8. Bitrate",
          "9. RTSP security",
      ],
      "What to verify": [
          "Ping device IP / check local routing",
          "Usually 554; confirm actual port",
          "Username/password and permissions",
          "Channel and stream number/path",
          "Special characters in credentials",
          "H.264 vs H.265/HEVC",
          "Very high FPS or long GOP",
          "High VBR/very high bitrate",
          "RTSP over TLS/RTSPS or encryption",
      ],
      "Typical action": [
          "Correct VLAN/routing or ask client IT to permit access.",
          "Use the actual configured RTSP port.",
          "Use a valid RTSP-capable account.",
          "Use exact vendor/model path.",
          "URL-encode reserved characters.",
          "Prefer H.264 for compatibility.",
          "Use ~15 FPS and GOP 15 as baseline.",
          "Use CBR and tune bitrate.",
          (
              "Use standard RTSP if the integration requires it; follow"
              " approved security policy."
          ),
      ],
  }
  st.table(pd.DataFrame(trouble_seq_data))

  # 12. CP Plus – Short Fix When Standard RTSP Fails
  st.markdown("### 12. CP Plus – Short Fix When Standard RTSP Fails")
  st.write(
      "CP Plus menus and capabilities vary by model/firmware. Do not assume"
      " the same menu exists on every device. For models exposing RTSP over"
      " TLS, check the RTSP-over-TLS setting. If the Edge Relay integration"
      " expects standard RTSP and the device is configured only for RTSPS/TLS,"
      " the stream may fail until the compatible mode is enabled."
  )
  st.write(
      "If ONVIF authentication is causing a separate discovery/configuration"
      " problem, review the ONVIF authentication setting only with the client's"
      " approval. ONVIF is not the same thing as RTSP authentication;"
      " disabling ONVIF authentication is not a universal fix for an RTSP"
      " failure."
  )

  # 13. What Can Happen If Recommended Settings Are Not Applied?
  st.markdown(
      "### 13. What Can Happen If Recommended Settings Are Not Applied?"
  )
  settings_impact_data = {
      "Setting not aligned": [
          "H.265/HEVC instead of H.264",
          "FPS much higher than required",
          "GOP much larger than FPS",
          "VBR with high peaks",
          "Bitrate too high",
          "RTSP over TLS/RTSPS when unsupported",
          "Wrong channel/stream number",
          "Wrong username/password",
          "Changing IP unexpectedly",
          "UDP on unstable network",
      ],
      "Possible effect": [
          (
              "VLC may work but Edge/cloud pipeline may fail or require extra"
              " decoder support"
          ),
          (
              "Higher bandwidth/CPU/GPU load and more frames to process"
          ),
          "Slower recovery after packet loss/reconnect",
          "Unexpected bandwidth spikes and buffering",
          "Latency, buffering, dropped frames",
          "Connection/authentication failure",
          "Black screen or wrong camera feed",
          "401/Unauthorized or connection failure",
          "Stream becomes unreachable",
          "Packet loss/artifacts/stalls",
      ],
      "Reason": [
          "Codec compatibility depends on the media pipeline.",
          "More frames per second increase processing and network load.",
          "Keyframes are farther apart.",
          "Peak bitrate can exceed planned capacity.",
          "Network and decoder load increase.",
          "Client and relay must support the same transport/security mode.",
          "RTSP path points to another channel/stream.",
          "Device rejects authentication.",
          "DHCP/address change breaks the source URL.",
          "UDP has less transport-level recovery than TCP.",
      ],
  }
  st.table(pd.DataFrame(settings_impact_data))

  # 14. RTMP Output – Client ID and Camera ID
  st.markdown("### 14. RTMP Output – Client ID and Camera ID")
  st.write("Anvex provides the Client ID.")
  st.code(
      "Template:\nrtmp://vision-media-server.anvex.ai:1935/{Client_ID}/{Camera_ID}",
      language="text",
  )
  st.write("Example Client ID: `6a8d6514a73283244b121f79`")
  st.write(
      "- **Camera 01**:"
      " `rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam01`"
  )
  st.write(
      "- **Camera 02**:"
      " `rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam02`"
  )
  st.caption(
      "Only the Client ID and Camera ID change. The media-server hostname and"
      " port remain unchanged unless Anvex infrastructure provides a different"
      " target."
  )

  # 16. streams.yaml – Final Configuration
  st.markdown("### 16. streams.yaml – Final Configuration")
  st.write(
      "Create/update `edge-relay/config/streams.yaml`. The file should contain"
      " only validated sources and the corresponding Anvex RTMP targets."
  )
  sample_yaml = """server:
  listen: 127.0.0.1:8080
  log_level: info

defaults:
  rtsp_transport: tcp
  srt_latency_ms: 200
  srt_connect_timeout_ms: 3000
  srt_payload_size: 1316
  on_stall_restart_s: 15
  startup_timeout_s: 10
  restart_backoff_base_s: 1
  restart_backoff_max_s: 60
  healthy_reset_s: 60
  srt_failure_threshold: 3
  rtmp_hold_s: 60
  rtmp_hold_max_s: 600
  kill_grace_s: 5

streams:
  - id: cam01
    source: rtsp://USER:ENCODED_PASSWORD@192.168.1.101:554/Streaming/Channels/101
    transport: rtmp
    target: rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam01
    fallback: rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam01
    srt_streamid: cam01

  - id: cam02
    source: rtsp://USER:ENCODED_PASSWORD@192.168.1.102:554/Streaming/Channels/101
    transport: rtmp
    target: rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam02
    fallback: rtmp://vision-media-server.anvex.ai:1935/6a8d6514a73283244b121f79/cam02
    srt_streamid: cam02"""
  st.code(sample_yaml, language="yaml")
  st.warning(
      "**Security**: Avoid committing real passwords into source control. If"
      " the deployment architecture supports secrets/environment variables,"
      " use the approved secret-management method."
  )

  # 17. Ubuntu / Docker Pre-Deployment
  st.markdown("### 17. Ubuntu / Docker Pre-Deployment")
  docker_script = """# Verify OS
cat /etc/os-release

# Verify Docker
docker --version
docker compose version

# If Docker is not installed
sudo apt update && sudo apt install -y ca-certificates curl
sudo install -m 0755 -d /etc/apt/keyrings
sudo curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
sudo chmod a+r /etc/apt/keyrings/docker.asc
echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu $(. /etc/os-release && echo "$VERSION_CODENAME") stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
sudo systemctl enable --now docker
docker --version"""
  st.code(docker_script, language="bash")

  # 18. Edge Relay Deployment
  st.markdown("### 18. Edge Relay Deployment")
  st.markdown("""
    1. Create the deployment directory: `mkdir -p ~/Desktop/anvex`
    2. Copy/unpack the edge-relay directory into `~/Desktop/anvex/`.
    3. Place the validated `streams.yaml` at `~/Desktop/anvex/edge-relay/config/streams.yaml`.
    4. Run the `deploy.sh` or `install.sh`.
    5. Review container/service logs and confirm all expected camera IDs start successfully.
    """)

  # 19. Post-Deployment Verification
  st.markdown("### 19. Post-Deployment Verification")
  st.markdown("""
    1. Open the Edge Relay dashboard: `http://<server-ip>:8080/streams`
    2. Confirm each Camera ID is present.
    3. Confirm the source is receiving bytes and the stream is not repeatedly restarting.
    4. Confirm RTMP output is reaching the Anvex media server.
    5. Observe the stream for several minutes for stalls, reconnects, and unexpected latency.
    6. Update the Excel inventory with **Deployment Status = Live** and record any remarks.
    """)

  # 20. Troubleshooting Matrix
  st.markdown("### 20. Troubleshooting Matrix")
  trouble_matrix_data = {
      "Symptom": [
          "VLC cannot open stream",
          "401 Unauthorized",
          "404 / path not found",
          "Black screen",
          "Video works in VLC but Edge Relay fails",
          "Frequent restarts",
          "Dashboard shows Connecting",
          "Wrong camera appears",
          "RTMP target unavailable",
          "YAML rejected",
      ],
      "Likely Cause": [
          "Wrong URL, credentials, port, path, or network",
          "Invalid credentials / account permission",
          "Wrong RTSP path",
          "Unsupported codec/profile or wrong stream",
          "Codec/transport/URL parsing difference",
          "Network loss, high bitrate, stalls",
          "Source unavailable or configuration error",
          "Incorrect channel mapping",
          "DNS/firewall/server issue",
          "Indentation/syntax/invalid field",
      ],
      "Action": [
          "Validate IP/port, credentials, URL path, and device access.",
          "Verify username/password and RTSP permission.",
          "Verify exact vendor/model/channel path.",
          "Use H.264 and verify main/sub stream path.",
          (
              "Check H.264, TCP transport, password encoding, and relay"
              " logs."
          ),
          "Use TCP, tune bitrate, verify bandwidth and camera stability.",
          "Test same RTSP URL with VLC from the deployment machine.",
          "Correct NVR/DVR channel and Camera ID mapping.",
          "Check outbound network/DNS and Anvex media server status.",
          "Validate YAML and compare against approved schema.",
      ],
  }
  st.table(pd.DataFrame(trouble_matrix_data))

  # 21. Pre-Deployment Checklist
  st.markdown("### 21. Pre-Deployment Checklist")
  checklist = [
      "Client ID received and verified.",
      "All agreed camera locations listed.",
      "Camera/NVR/DVR IP addresses received.",
      "RTSP port confirmed.",
      "RTSP username/password received through approved secure method.",
      "NVR/DVR credentials received if configuration access is required.",
      "Channel numbers mapped to physical camera names.",
      "Main RTSP URL identified.",
      "Sub RTSP URL identified where required.",
      "Every required RTSP feed tested successfully in VLC.",
      "H.264 / FPS / GOP / bitrate / CBR settings checked.",
      "CP Plus / vendor-specific settings checked where applicable.",
      "Camera IDs assigned and frozen.",
      "RTMP URLs generated from Client ID + Camera ID.",
      "streams.yaml reviewed and validated.",
      "Docker/Edge Relay host ready.",
      "Network/firewall/DNS connectivity checked.",
  ]
  for item in checklist:
    st.checkbox(item, key=f"chk_{hash(item)}")