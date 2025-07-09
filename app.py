import streamlit as st
import streamlit_authenticator as stauth
import requests
import openpyxl
import json
import re

names = ['Chris', 'Steve']
usernames = ['bochtisc@gmail.com', 'm-oconsultancy@outlook.com']
hashed_passwords = [
    '$2b$12$2txxtWw8smMumw6J8R0v1.yPL3TS1k4tC6TOkAlq9UJJcCCJhB8rG',   # password: 123456789
    '$2b$12$MRKfjsYOdcvyyD7s6vZubecf5k6bHeoAM9B79vBC.4Pa8ljAvKLCa'    # password: 12345678910
]

authenticator = stauth.Authenticate(
    names, usernames, hashed_passwords,
    'cookie_name', 'signature_key', cookie_expiry_days=1
)

name, authentication_status, username = authenticator.login('Σύνδεση', 'main')

if authentication_status is None:
    st.warning('Συμπλήρωσε email και κωδικό για να μπεις.')
    st.stop()
elif authentication_status is False:
    st.error('Λάθος email ή κωδικός.')
    st.stop()
elif authentication_status:
    st.success(f'Καλωσήρθες {name}!')


# ======= ΤΟ ΚΥΡΙΩΣ APP ΞΕΚΙΝΑ ΑΠΟ ΕΔΩ ===========

st.markdown("""
    <style>
    .stDownloadButton button {
        background-color: #0066cc !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

API_TOKEN = "pk_82763580_PX00W04XWNJPJ2YR4M6NCNZ8WQPOLY6O"
LIST_ID = "901511575020"
HEADERS = {"Authorization": API_TOKEN}

with open("fields_mapping.json", "r", encoding="utf-8") as f:
    CLICKUP_FIELDS = json.load(f)

def get_tasks(list_id):
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task?archived=false"
    resp = requests.get(url, headers=HEADERS)
    if resp.status_code != 200:
        st.error(f"ClickUp API Error: {resp.status_code} - {resp.text}")
        return []
    return resp.json().get("tasks", [])

def find_task_by_name(tasks, name):
    for task in tasks:
        if task.get("name") == name:
            return task
    return None

def extract_custom_fields(task):
    out = {}
    for field in task.get("custom_fields", []):
        name = field.get("name")
        value = field.get("value")
        if value is not None:
            out[name] = str(value)
    return out

def clean_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        parts = value.strip().split()
        num = parts[0].replace(",", ".")
        return num
    return str(value)

def patch_or_insert_tag(xml_text, mapping, values):
    new_xml = xml_text
    for field in mapping:
        if field.get("source") not in ("clickup", "excel", "custom"):
            continue
        value = values.get(field["field"], "")
        xml_path = field.get("xml_path")
        if not xml_path or '/' not in xml_path:
            continue
        path_parts = xml_path.strip('./').split('/')
        tag = path_parts[-1]
        parent_tag = path_parts[-2] if len(path_parts) > 1 else None
        if parent_tag:
            parent_pattern = re.compile(
                f'(<{parent_tag}.*?>)(.*?)(</{parent_tag}>)',
                re.DOTALL | re.IGNORECASE
            )
            def replace_in_parent(m):
                content = m.group(2)
                tag_pattern = re.compile(f'(<{tag}>)(.*?)(</{tag}>)', re.DOTALL)
                if tag_pattern.search(content):
                    new_content = tag_pattern.sub(rf'\1{value}\3', content)
                    return m.group(1) + new_content + m.group(3)
                else:
                    insert_text = f'<{tag}>{value}</{tag}>'
                    new_content = content + insert_text
                    return m.group(1) + new_content + m.group(3)
            new_xml = parent_pattern.sub(replace_in_parent, new_xml)
        else:
            tag_pattern = re.compile(f'(<{tag}>)(.*?)(</{tag}>)', re.DOTALL)
            new_xml = tag_pattern.sub(rf'\1{value}\3', new_xml)
    return new_xml

st.title("XML Update")

uploaded_files = st.file_uploader(
    "Ανέβασε τα δύο αρχεία (XML + Excel με ίδιο όνομα, drag & drop μαζί)",
    type=["xml", "xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and len(uploaded_files) >= 2:
    files_by_type = {}
    for f in uploaded_files:
        ext = f.name.rsplit('.', 1)[-1].lower()
        base = f.name.rsplit('.', 1)[0]
        files_by_type.setdefault(ext, {})[base] = f

    xml_files = files_by_type.get('xml', {})
    excel_files = {**files_by_type.get('xlsx', {}), **files_by_type.get('xls', {})}
    common_names = set(xml_files) & set(excel_files)

    if not common_names:
        st.error("Δε βρέθηκε XML & Excel με ίδιο όνομα!")
    else:
        selected = list(common_names)[0]
        xml_file = xml_files[selected]
        excel_file = excel_files[selected]
        try:
            xml_text = xml_file.read().decode("utf-8")
        except Exception as e:
            st.error(f"Σφάλμα στο διάβασμα του XML: {e}")
            st.stop()

        # Αυτόματη προσθήκη <Verdieping> αν λείπει
        if "<Verdiepingen>" in xml_text and "<Verdieping>" not in xml_text:
            xml_text = xml_text.replace("<Verdiepingen>", "<Verdiepingen><Verdieping></Verdieping>")

        # Ενημέρωση από Excel
        excel_status = st.empty()
        excel_status.info(f"Ενημέρωση XML από Excel ({selected})...")
        excel_values = {}
        try:
            excel_file.seek(0)
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            if "Algemeen" not in wb.sheetnames:
                excel_status.error("❌ Το Excel δεν περιέχει φύλλο με όνομα 'Algemeen'!")
                st.stop()
            ws = wb["Algemeen"]
            for field in CLICKUP_FIELDS:
                if field.get("source") == "excel":
                    cell = field.get("cell", "").replace(" ", "")
                    value = ws[cell].value if cell else ""
                    if field["field"] == "Gebruiksoppervlakte" and (value is None or value == ""):
                        value = "0"
                    excel_values[field["field"]] = clean_excel_value(value)
        except Exception as e:
            excel_status.error(f"❌ Σφάλμα κατά το διάβασμα του Excel: {e}")
            st.stop()

        missing_excels = [f["ui_label"] for f in CLICKUP_FIELDS if f.get("source") == "excel" and not excel_values.get(f["field"])]
        if missing_excels:
            excel_status.warning("Κενά πεδία Excel: " + ", ".join(missing_excels))
        else:
            excel_status.success("Ενημέρωση XML από Excel")

        for field in CLICKUP_FIELDS:
            if field.get("source") == "excel":
                val = excel_values.get(field["field"], "")
                st.markdown(f"✅ Τιμή {field['ui_label']} από Excel ({field.get('cell','')}): <b>{val}</b>", unsafe_allow_html=True)

        # ClickUp data
        clickup_status = st.empty()
        clickup_status.info(f"Ψάχνω στο ClickUp για task με όνομα: {selected}")
        tasks = get_tasks(LIST_ID)
        if not tasks:
            clickup_status.error("❌ Δε βρέθηκαν tasks στο ClickUp ή αποτυχία σύνδεσης.")
            st.stop()
        task = find_task_by_name(tasks, selected)
        if not task:
            clickup_status.error(f"❌ Δε βρέθηκε task στο ClickUp με όνομα '{selected}'.")
            st.stop()
        fields = extract_custom_fields(task)
        clickup_status.success(f"Βρέθηκε task στο ClickUp: {selected}")

        xml_status = st.empty()
        col1, col2 = st.columns([8, 1])
        with col2:
            edit_mode = st.toggle("✏️", key="edit_fields")
        with col1:
            st.markdown("### Συμπλήρωση στο XML από ClickUp & Excel & Custom πεδία (με fixed value)")

        updated_fields = {}

        # ClickUp πεδία
        for field in CLICKUP_FIELDS:
            if field.get("source") == "clickup":
                ck_field = field["field"]
                xml_label = field["ui_label"]
                value = fields.get(ck_field, "")
                if edit_mode:
                    updated_value = st.text_input(f"{xml_label}", value=value, key=f"edit_{ck_field}")
                else:
                    icon = "✅" if value else "❌"
                    st.markdown(f"{icon} <b>{xml_label}</b>: <span style='color:#222'>{value}</span>", unsafe_allow_html=True)
                    updated_value = value
                updated_fields[ck_field] = updated_value.strip() if updated_value else ""

        # Excel πεδία
        for field in CLICKUP_FIELDS:
            if field.get("source") == "excel":
                updated_fields[field["field"]] = excel_values.get(field["field"], "")

        # Custom/fixed value πεδία από το mapping (αν υπάρχει fixed_value)
        for field in CLICKUP_FIELDS:
            if "fixed_value" in field:
                updated_fields[field["field"]] = field["fixed_value"]

        missing_fields = [
            field["ui_label"]
            for field in CLICKUP_FIELDS
            if field.get("source") in ("clickup", "excel", "custom") and not updated_fields.get(field["field"])
        ]
        if missing_fields:
            st.warning("Λείπουν πεδία: " + ", ".join(missing_fields) +
                       " — Τα αντίστοιχα πεδία στο XML θα μείνουν κενά.")
        else:
            xml_status.empty()

        # ΕΦΑΡΜΟΓΗ PATCH στο XML
        try:
            new_xml = patch_or_insert_tag(xml_text, CLICKUP_FIELDS, updated_fields)
            st.download_button(
                label="Κατέβασε το νέο XML",
                data=new_xml,
                file_name=f"{selected}!.xml",
                mime="application/xml"
            )
        except Exception as e:
            xml_status.error(f"Σφάλμα: {e}")

else:
    st.info("Ανέβασε δύο αρχεία με το ίδιο όνομα (XML & Excel) μαζί, με drag & drop.")
