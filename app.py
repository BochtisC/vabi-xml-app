import streamlit as st
import requests
import openpyxl
import json
import re
import pandas as pd
import datetime

st.markdown("""
    <style>
    .stDownloadButton button {
        background-color: #0066cc !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

API_TOKEN = "pk_82763580_PX00W04XWNJPJ2YR4M6NCNZ8WQPOLY6O"
LIST_IDS = [
    "901206264874",      # Energielabel Haaglanden
    "901511575020",      # Mijn EnergieLabel
    "901504459596",      # Energieinspectie
]

LIST_MAPPINGS = {
    "901206264874": "mapping-Energielabel_Haaglanden.json",
    "901511575020": "mapping-Mijn_EnergieLabel.json",
    "901504459596": "mapping-EnergieInspectie.json",
}

LIST_NAMES = {
    "901206264874": "Energielabel Haaglanden",
    "901511575020": "Mijn EnergieLabel",
    "901504459596": "Εnergie Inspectie",
}

HEADERS = {"Authorization": API_TOKEN}

def get_tasks(list_id):
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task?archived=false"
    resp = requests.get(url, headers=HEADERS)
    if resp.status_code != 200:
        return []
    return resp.json().get("tasks", [])

def extract_custom_fields(task):
    out = {}
    dropdown_options = {}
    for f in task.get("custom_fields", []):
        if f.get("type") in ["drop_down", "dropdown"] and "type_config" in f and "options" in f["type_config"]:
            dropdown_options[f["name"]] = f["type_config"]["options"]
    for field in task.get("custom_fields", []):
        name = field.get("name")
        value = field.get("value")
        typ = field.get("type")
        if typ in ["drop_down", "dropdown"] and value is not None:
            options = dropdown_options.get(name, [])
            label = None
            for opt in options:
                if ("orderindex" in opt and opt["orderindex"] == value) or ("id" in opt and opt["id"] == value):
                    label = opt["name"]
                    break
            # STRIP PATCH: strip label if string
            if label is not None and isinstance(label, str):
                label = label.strip()
            out[name] = {"id": value, "label": label if label is not None else str(value).strip()}
        elif value is not None:
            # STRIP PATCH: strip string values!
            out[name] = str(value).strip() if isinstance(value, str) else str(value)
        else:
            out[name] = ""
    return out

def clean_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        parts = value.strip().split()
        num = parts[0].replace(",", ".")
        return num.strip()     # STRIP PATCH: extra strip
    return str(value).strip()  # STRIP PATCH

def get_rz_values(ws, column, num_rows):
    vals = []
    for i in range(num_rows):
        row = 3 + i
        v = ws[f"{column}{row}"].value
        if v is not None and str(v).strip() != "":
            num = str(v).split()[0].replace(",", ".")
            vals.append(num.strip())    # STRIP PATCH
        else:
            vals.append("0")
    return vals

def update_verdiepingen_in_rekenzone(xml_text, rz_name, verdieping_values):
    pattern = re.compile(
        rf'(<Rekenzone>.*?<Naam>{rz_name}</Naam>.*?<Verdiepingen.*?>)(.*?)(</Verdiepingen>)(.*?</Rekenzone>)',
        re.DOTALL)
    def replacement(match):
        head, old_content, tail, after = match.group(1), match.group(2), match.group(3), match.group(4)
        new_content = ""
        for val in verdieping_values:
            new_content += f"<Verdieping><Gebruiksoppervlakte>{val}</Gebruiksoppervlakte></Verdieping>"
        return head + new_content + tail + after
    return pattern.sub(replacement, xml_text)

def checkmark(val):
    try:
        return f"✅ {val}" if float(val) != 0 else f"❌ 0"
    except:
        return val

def smart_patch_xml(xml_text, mappings, values, root_tag='Objecten'):
    xml = xml_text

    def patch_xml_tag(xml, xml_path, value):
        path_parts = xml_path.strip('./').split('/')
        if not path_parts:
            return xml
        parent_path = path_parts[:-1]
        final_tag = path_parts[-1]
        parent_tag = parent_path[-1] if parent_path else None

        new_tag = f"<{final_tag}>{value}</{final_tag}>"

        if parent_tag:
            parent_pattern = rf'(<{parent_tag}[^>]*>)(.*?)(</{parent_tag}>)'
            m = re.search(parent_pattern, xml, flags=re.DOTALL)
            if m:
                inside = m.group(2)
                tag_pattern = rf'(<{final_tag}>)(.*?)(</{final_tag}>)'
                if re.search(tag_pattern, inside, flags=re.DOTALL):
                    new_inside = re.sub(tag_pattern, rf'\1{value}\3', inside, flags=re.DOTALL)
                else:
                    new_inside = inside + new_tag
                return xml[:m.start(2)] + new_inside + xml[m.end(2):]
            else:
                object_algemeen_pattern = r'(<ObjectAlgemeen[^>]*>)(.*?)(</ObjectAlgemeen>)'
                obj_match = re.search(object_algemeen_pattern, xml, flags=re.DOTALL)
                if obj_match:
                    parent_block = f"<{parent_tag}>{new_tag}</{parent_tag}>"
                    new_content = parent_block + obj_match.group(2)
                    return (
                        xml[:obj_match.start(2)] +
                        new_content +
                        xml[obj_match.end(2):]
                    )
                else:
                    root_tag = path_parts[0]
                    root_open = re.search(rf'(<{root_tag}[^>]*>)', xml, flags=re.IGNORECASE)
                    if root_open:
                        parent_block = f"<{parent_tag}>{new_tag}</{parent_tag}>"
                        return xml[:root_open.end()] + parent_block + xml[root_open.end():]
                    else:
                        return f"<{parent_tag}>{new_tag}</{parent_tag}>" + xml
        else:
            root_tag = path_parts[0]
            root_open = re.search(rf'(<{root_tag}[^>]*>)', xml, flags=re.IGNORECASE)
            if root_open:
                return xml[:root_open.end()] + new_tag + xml[root_open.end():]
            else:
                return new_tag + xml

    for field in mappings:
        xml_path = field.get("xml_path")
        if not xml_path or "/" not in xml_path:
            continue
        value = values.get(field["field"], "")
        if isinstance(value, dict):
            xml_value_type = field.get("xml_value_type", "id")
            value = value.get(xml_value_type, value.get("id", ""))
        # STRIP PATCH
        if isinstance(value, str):
            value = value.strip()
        if value != "":
            xml = patch_xml_tag(xml, xml_path, value)
    return xml

# ---- END OF PATCHES ----

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

        excel_file.seek(0)
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        if "Algemeen" not in wb.sheetnames:
            st.error("❌ Το Excel δεν περιέχει φύλλο με όνομα 'Algemeen'!")
            st.stop()
        ws = wb["Algemeen"]

        rz1_vals = []
        for i in range(3, 13):
            v = ws[f"B{i}"].value
            if v is not None and str(v).strip() != "":
                num = str(v).split()[0].replace(",", ".")
                rz1_vals.append(num.strip())   # STRIP PATCH
        num_verdiepingen = len(rz1_vals)
        rz2_vals = get_rz_values(ws, "C", num_verdiepingen)
        rz3_vals = get_rz_values(ws, "D", num_verdiepingen)

        data = []
        for i in range(num_verdiepingen):
            row = {
                "Verdieping": f"Verdieping {i+1} (B{3+i}/C{3+i}/D{3+i})",
                "rz1 (B)": checkmark(rz1_vals[i]),
                "rz2 (C)": checkmark(rz2_vals[i]),
                "rz3 (D)": checkmark(rz3_vals[i])
            }
            data.append(row)
        sum_b = sum([float(x) for x in rz1_vals])
        sum_c = sum([float(x) for x in rz2_vals])
        sum_d = sum([float(x) for x in rz3_vals])
        data.append({
            "Verdieping": "ΣΥΝΟΛΟ",
            "rz1 (B)": f"{sum_b:.2f}",
            "rz2 (C)": f"{sum_c:.2f}",
            "rz3 (D)": f"{sum_d:.2f}"
        })
        df = pd.DataFrame(data)

        st.success(f"Βρέθηκε Excel: {selected}.xlsx (φύλλο: Algemeen)")
        st.markdown("#### Τιμές Verdiepingen ανά rz (από Excel):")
        st.table(df)

        task = None
        found_list_id = None
        for list_id in LIST_IDS:
            try:
                tasks = get_tasks(list_id)
            except Exception:
                continue
            for t in tasks:
                if t.get("name") == selected:
                    task = t
                    found_list_id = list_id
                    break
            if task:
                break

        if not task:
            st.error(f"❌ Δε βρέθηκε task στο ClickUp με όνομα '{selected}'.")
            st.stop()

        mapping_file = LIST_MAPPINGS.get(found_list_id)
        if mapping_file is None:
            st.error(f"❌ Δεν βρέθηκε mapping για λίστα: {found_list_id}")
            st.stop()
        with open(mapping_file, "r", encoding="utf-8") as f:
            CLICKUP_FIELDS = json.load(f)

        excel_values = {}
        for field in CLICKUP_FIELDS:
            if field.get("source") == "excel":
                cell = field.get("cell", "").replace(" ", "")
                value = ws[cell].value if cell else ""
                if field["field"] == "Gebruiksoppervlakte" and (value is None or value == ""):
                    value = "1"
                excel_values[field["field"]] = clean_excel_value(value)

        excel_block = ""
        for field in CLICKUP_FIELDS:
            if field.get("source") == "excel":
                val = excel_values.get(field["field"], "")
                excel_block += f"✅ Τιμή <b>{field['ui_label']}</b> από Excel (<b>{field.get('cell','')}</b>): <b>{val}</b><br>"
        if excel_block:
            st.markdown(excel_block, unsafe_allow_html=True)

        clickup_status = st.empty()
        clickup_status.info(f"Ψάχνω στο ClickUp για task με όνομα: {selected}")

        st.success(f"Βρέθηκε task στο ClickUp: {selected} στη λίστα: {LIST_NAMES.get(found_list_id, found_list_id)}")

        fields = extract_custom_fields(task)

        if "date_created" in task:
            try:
                ts = int(task["date_created"]) / 1000
                fields["date_created"] = datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
            except Exception:
                fields["date_created"] = ""

        for field in CLICKUP_FIELDS:
            if field.get("source") == "clickup":
                val = fields.get(field["field"], "")
                if field["field"] == "date_created":
                    label = val
                else:
                    label = val["label"] if isinstance(val, dict) else val
                icon = "✅" if label else "❌"
                st.markdown(f"{icon} <b>{field['ui_label']}</b>: <span style='color:#222'>{label or '(κενό)'}</span><br>", unsafe_allow_html=True)
            elif field.get("source") == "excel":
                val = excel_values.get(field["field"], "")
                st.markdown(f"📄 <b>{field['ui_label']}</b>: <span style='color:#222'>{val}</span>", unsafe_allow_html=True)
            elif field.get("source") == "custom":
                val = field.get("fixed_value", "")
                st.markdown(f"🔒 <b>{field['ui_label']}</b>: <span style='color:#222'>{val}</span>", unsafe_allow_html=True)

        updated_fields = {}
        for field in CLICKUP_FIELDS:
            if field.get("source") == "clickup":
                val = fields.get(field["field"], "")
                if isinstance(val, dict):
                    updated_fields[field["field"]] = val
                elif isinstance(val, str):
                    updated_fields[field["field"]] = val.strip()     # STRIP PATCH
                else:
                    updated_fields[field["field"]] = val
            elif field.get("source") == "excel":
                updated_fields[field["field"]] = excel_values.get(field["field"], "")
            elif field.get("source") == "custom":
                updated_fields[field["field"]] = field.get("fixed_value", "")

        # -- STRIP PATCH: strip όλα τα string values στα updated_fields --
        for k, v in updated_fields.items():
            if isinstance(v, str):
                updated_fields[k] = v.strip()

        try:
            new_xml = smart_patch_xml(xml_text, CLICKUP_FIELDS, updated_fields)
            if re.search(r'<Rekenzone>.*?<Naam>rz1</Naam>', new_xml, re.DOTALL):
                new_xml = update_verdiepingen_in_rekenzone(new_xml, "rz1", rz1_vals)
            if re.search(r'<Rekenzone>.*?<Naam>rz2</Naam>', new_xml, re.DOTALL):
                new_xml = update_verdiepingen_in_rekenzone(new_xml, "rz2", rz2_vals)
            if re.search(r'<Rekenzone>.*?<Naam>rz3</Naam>', new_xml, re.DOTALL):
                new_xml = update_verdiepingen_in_rekenzone(new_xml, "rz3", rz3_vals)

            st.download_button(
                label="Κατέβασε το νέο XML",
                data=new_xml,
                file_name=f"{selected}!.xml",
                mime="application/xml"
            )
        except Exception as e:
            st.error(f"Σφάλμα: {e}")

else:
    st.info("Ανέβασε δύο αρχεία με το ίδιο όνομα (XML & Excel) μαζί, με drag & drop.")
