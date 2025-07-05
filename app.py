import streamlit as st
import requests
import openpyxl
import re

# ---- ClickUp Settings ----
API_TOKEN = "pk_82763580_PX00W04XWNJPJ2YR4M6NCNZ8WQPOLY6O"
LIST_ID = "901511575020"
HEADERS = {"Authorization": API_TOKEN}

CLICKUP_FIELDS = [
    ("A1c Straat", "Straat"),
    ("A1c Huisnummer", "Huisnummer"),
    ("A1c huisnummer toev.", "Huisnummertoevoeging"),
    ("A1c Postcode", "Postcode"),
    ("A1c Plaats", "Plaats"),
    ("A1c Adres", "NaamObject"),  # Αυτό είναι για το <NaamObject>
]

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

def safe_patch_object_adresgegevens(xml_text, fields):
    new_adres = (
        "<ObjectAdresgegevens>"
        f"<Straat>{fields.get('A1c Straat', '')}</Straat>"
        f"<Huisnummer>{fields.get('A1c Huisnummer', '')}</Huisnummer>"
        f"<Huisletter>{fields.get('A1c Huisletter', '')}</Huisletter>"
        f"<Huisnummertoevoeging>{fields.get('A1c huisnummer toev.', '')}</Huisnummertoevoeging>"
        f"<Postcode>{fields.get('A1c Postcode', '')}</Postcode>"
        f"<Woonplaats>{fields.get('A1c Plaats', '')}</Woonplaats>"
        f"<BagId></BagId>"
        "</ObjectAdresgegevens>"
    )
    old_tag = "<ObjectAdresgegevens></ObjectAdresgegevens>"
    if old_tag not in xml_text:
        st.error(f"Δεν βρέθηκε το tag {old_tag} στο XML!")
        return xml_text
    return xml_text.replace(old_tag, new_adres, 1)

def safe_patch_object_classificatie(xml_text, gebouwhoogte):
    old_tag = "<ObjectClassificatie></ObjectClassificatie>"
    new_tag = (
        "<ObjectClassificatie>"
        f"<Gebouwhoogte>{gebouwhoogte}</Gebouwhoogte>"
        "</ObjectClassificatie>"
    )
    if old_tag in xml_text:
        return xml_text.replace(old_tag, new_tag, 1)
    else:
        if "</EPA>" in xml_text:
            return xml_text.replace("</EPA>", new_tag + "\n</EPA>")
        else:
            return xml_text + "\n" + new_tag

def safe_patch_object_naamobject(xml_text, naamobject):
    pattern = r"(<NaamObject>)(.*?)(</NaamObject>)"
    new_text = re.sub(pattern, r"\1{}\3".format(naamobject), xml_text, count=1, flags=re.DOTALL)
    if new_text == xml_text:
        st.error("Δεν βρέθηκε το tag <NaamObject> στο XML!")
    return new_text

# --- UI ---
st.title("Vabi XML ")

uploaded_files = st.file_uploader(
    "Ανέβασε τα δύο αρχεία (XML + Excel με ίδιο όνομα, drag & drop μαζί)",
    type=["xml", "xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and len(uploaded_files) >= 2:
    # Ταξινόμηση αρχείων
    files_by_type = {}
    for f in uploaded_files:
        ext = f.name.rsplit('.', 1)[-1].lower()
        base = f.name.rsplit('.', 1)[0]
        files_by_type.setdefault(ext, {})[base] = f

    xml_files = files_by_type.get('xml', {})
    excel_files = {**files_by_type.get('xlsx', {}), **files_by_type.get('xls', {})}
    common_names = set(xml_files) & set(excel_files)

    if not common_names:
        st.error("Δε βρέθηκε ζευγάρι XML/Excel με ίδιο όνομα (χωρίς κατάληξη)!")
    else:
        selected = list(common_names)[0]
        st.info(f"Βρέθηκε ζευγάρι: {selected}.xml & {selected}.xlsx/xls")
        xml_file = xml_files[selected]
        excel_file = excel_files[selected]

        xml_text = xml_file.read().decode("utf-8")

        # --- Excel: Gebouwhoogte ---
        try:
            excel_file.seek(0)
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            if "Algemeen" not in wb.sheetnames:
                st.error("Το Excel δεν περιέχει φύλλο με όνομα 'Algemeen'!")
                st.stop()
            ws = wb["Algemeen"]
            gebouwhoogte = ws["N6"].value
        except Exception as e:
            st.error(f"Σφάλμα κατά το διάβασμα του Excel: {e}")
            st.stop()

        if gebouwhoogte is None or str(gebouwhoogte).strip() == "":
            st.error("Το κελί N6 του φύλλου 'Algemeen' είναι κενό!")
            st.stop()
        st.markdown(f"Τιμή Gebouwhoogte από Excel (Algemeen!N6): **{gebouwhoogte}**")

        # --- ClickUp ---
        st.info(f"Ψάχνω στο ClickUp για task με όνομα: **{selected}**")
        tasks = get_tasks(LIST_ID)
        if not tasks:
            st.error("Δε βρέθηκαν tasks στο ClickUp ή αποτυχία σύνδεσης.")
            st.stop()
        task = find_task_by_name(tasks, selected)
        if not task:
            st.error(f"Δε βρέθηκε task στο ClickUp με όνομα '{selected}'.")
            st.stop()
        fields = extract_custom_fields(task)

        st.markdown("### Συμπλήρωση στο XML από Clickup")
        for ck_field, xml_label in CLICKUP_FIELDS:
            value = fields.get(ck_field, "")
            icon = "✅" if value else "❌"
            st.markdown(f"{icon} <b>{xml_label}</b>: <span style='color:#222'>{value}</span>", unsafe_allow_html=True)

        naamobject = fields.get("A1c Adres", "")

        if st.button("Ενημέρωση και Κατέβασμα XML"):
            try:
                # Πρώτα ενημερώνεις ObjectAdresgegevens
                new_xml = safe_patch_object_adresgegevens(xml_text, fields)
                # Μετά ObjectClassificatie (Excel)
                new_xml = safe_patch_object_classificatie(new_xml, gebouwhoogte)
                # Τέλος, NaamObject (ClickUp)
                new_xml = safe_patch_object_naamobject(new_xml, naamobject)
                st.success("Το XML ενημερώθηκε στα τρία σημεία!")
                st.download_button(
                    label="Κατέβασε το νέο XML",
                    data=new_xml,
                    file_name=f"{selected}_updated.xml",
                    mime="application/xml"
                )
                with st.expander("Δες το νέο XML"):
                    st.code(new_xml, language="xml")
            except Exception as e:
                st.error(f"Σφάλμα: {e}")

else:
    st.info("Ανέβασε δύο αρχεία με το ίδιο όνομα (XML & Excel) μαζί, με drag & drop.")
