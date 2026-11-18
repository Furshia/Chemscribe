import requests
import re
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from tkinter import Tk, filedialog, Checkbutton, IntVar, Button, Label
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from typing import Any
from pprint import pprint


pug_rest_base_url = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug'
pug_view_base_url = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug_view'
formula_pattern = r'(?:[A-Z][a-z]?\d*|\([A-Z][a-z]?\d*(?:[A-Z][a-z]?\d*)*\)\d*)+'




#formula_url = f'{pug_rest_base_url}/compound/fastformula/{encoded_formula}/cids/JSON?MaxRecords=100'
#rest_url = f'{pug_rest_base_url}/compound/cid/{cid_list}/property/IUPACName,MolecularWeight,ExactMass/JSON'
#rest_url = f'{pug_rest_base_url}/compound/name/{encoded_name}/property/MolecularWeight,ExactMass/JSON'
#view_url = f'{pug_view_base_url}/data/compound/{cid}/JSON'

def identify_phase(chemical: str) -> tuple[str, str]: #AI Function
    phase_match = re.search(
        r'\((gas|g|liquid|l|solid|s)\)',
        chemical,
        re.IGNORECASE
    )

    if phase_match:
        phase = {
            'g': 'gas',
            'gas': 'gas',
            'l': 'liquid',
            'liquid': 'liquid',
            's': 'solid',
            'solid': 'solid'
        }[phase_match.group(1).lower()]

        chemical = re.sub(
            r'\s*\((?:gas|g|liquid|l|solid|s)\)\s*',
            '',
            chemical,
            flags=re.IGNORECASE
        ).strip()

    else:
        phase = 'aqueous'

    return chemical, phase


def check_server() -> bool:  #check if pubchem is online
    url = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/water/property/MolecularWeight/JSON'
    try:
        response: any = requests.get(url, timeout=10)
        response.raise_for_status()
        return True
    except requests.exceptions.ConnectionError:
        print("No internet connection. Please check your network.")
        return False
    except requests.exceptions.RequestException as error:
        print(f"Request error: {error}")
        return False

def check_chemical_if_found(url: str) -> dict[str, dict[str, list[int]]] | None: #! Duplicated code with check_server, Something wrong with this function
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.json()

    except requests.exceptions.HTTPError as error:
        if error.response is not None and error.response.status_code == 404:
            return None
        else:
            print(f"PubChem HTTP error: {error.response.status_code if error.response else 'Unknown'}")
            return None

    except requests.exceptions.RequestException as error:
        print(f"Request error: {error}")
        return None
 
def chemical_not_found(chemical: str) -> None: #If chemical is not found #!Add later "Did you mean chemical (similar to user input found in pubchem)?"
    print(f'{chemical} wasn\'t found in the PubChem Database')






#Fetches data from pubchem, all are called inside the def_chemical_formula_fetch function #AI 
def fetch_optional_data(url: str) -> dict[str, Any] | None: # AI Function
    response = requests.get(url, timeout=10)

    if response.status_code == 404:
        return None

    response.raise_for_status()
    return response.json()
def find_ghs_symbols(sections: list[dict[str, Any]]) -> list[str]:# AI Function
    ghs_symbols: set[str] = set()

    for section in sections:
        if section.get('TOCHeading') == 'GHS Classification':

            for information in section.get('Information', []):
                if information.get('Name') != 'Pictogram(s)':
                    continue

                for item in information.get('Value', {}).get('StringWithMarkup', []):
                    for markup in item.get('Markup', []):
                        markup_text = (
                            markup.get('URL', '') + ' ' +
                            markup.get('Extra', '')
                        )

                        matches = re.findall(r'GHS\d{2}', markup_text)
                        ghs_symbols.update(matches)

        ghs_symbols.update(find_ghs_symbols(section.get('Section', [])))

    return sorted(ghs_symbols)
def is_valid_property_value(value: Any, heading: str) -> bool:  # AI Function

    if not isinstance(value, dict):
        return False

    entries = value.get('StringWithMarkup', [])

    for entry in entries:

        if not isinstance(entry, dict):
            continue

        text = entry.get('String')

        if not isinstance(text, str):
            continue

        text = text.strip()

        if heading in ('Melting Point', 'Boiling Point'):

            if re.search(
                r'[-+]?\d+(?:\.\d+)?\s*(?:°\s*[CF]|degrees?\s*[CF]|K\b)',
                text,
                re.IGNORECASE
            ):
                return True

        elif heading == 'Density':

            if re.match(
                r'^\s*[~<>≤≥]?\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)',
                text
            ):
                return True

    return False
def collect_information_records( # AI Function
    sections: list[dict[str, Any]],
    heading: str,
    references: list[dict[str, Any]] | None = None
) -> list[dict[str, Any]]: # AI Function

    records: list[dict[str, Any]] = []

    reference_lookup: dict[int, dict[str, Any]] = {}

    for reference in references or []:
        reference_number: Any = reference.get('ReferenceNumber')

        if isinstance(reference_number, int):
            reference_lookup[reference_number] = reference

    def collect_information(section: dict[str, Any]) -> None:

        section_heading: str = section.get('TOCHeading', heading)

        information_list: list[dict[str, Any]] = (
            section.get('Information', [])
        )

        for information in information_list:

            value_data: dict[str, Any] = information.get('Value', {})

            string_parts: list[dict[str, Any]] = (
                value_data.get('StringWithMarkup', [])
            )

            value: str = ' '.join(
                part['String']
                for part in string_parts
                if isinstance(part.get('String'), str)
            )

            if not value:
                continue

            reference_number: Any = information.get('ReferenceNumber')

            source: dict[str, Any] = (
                reference_lookup.get(reference_number, {})
                if isinstance(reference_number, int)
                else {}
            )

            year: int | None = extract_publication_year(information)

            if year is None:
                year = extract_publication_year(source)

            records.append({
                'name': information.get('Name') or section_heading,
                'value': value,
                'year': year,
                'reference_number': reference_number,
                'source_name': source.get('SourceName'),
                'source_id': source.get('SourceID'),
                'source_url': source.get('URL'),
                'reference': information.get('Reference'),
                'description': information.get('Description')
            })

        subsections: list[dict[str, Any]] = section.get('Section', [])

        for subsection in subsections:
            collect_information(subsection)

    def find_heading(current_sections: list[dict[str, Any]]) -> None:

        for section in current_sections:

            if section.get('TOCHeading') == heading:
                collect_information(section)

            else:
                find_heading(section.get('Section', []))

    def sort_key(record: dict[str, Any]) -> tuple[bool, int]:

        year: Any = record.get('year')

        return (
            not isinstance(year, int),
            -year if isinstance(year, int) else 0
        )

    find_heading(sections)

    records.sort(key=sort_key)

    return records
def collect_property_records(sections: list[dict[str, Any]], heading: str) -> list[dict[str, Any]]: # AI Function

    property_records: list[dict[str, Any]] = []

    for section in sections:

        if section.get('TOCHeading') == heading:

            for information in section.get('Information', []):

                value = information.get('Value')

                if not is_valid_property_value(value, heading):
                    continue

                property_records.append({
                    'value': value,
                    'reference': information.get('Reference'),
                    'year': extract_publication_year(information),
                    'reference_number': information.get('ReferenceNumber')
                })

        property_records.extend(
            collect_property_records(section.get('Section', []), heading)
        )

    return property_records
def extract_publication_year(information: dict[str, Any]) -> int | None: # AI Function

    year: Any = information.get('Year')

    if isinstance(year, int):
        return year

    if isinstance(year, str):
        match: re.Match[str] | None = re.search(
            r'(?<!\d)(?:18|19|20|21)\d{2}(?!\d)',
            year
        )

        if match:
            return int(match.group())

    text_values: list[str] = []

    for field in ('Name', 'Description', 'SourceName', 'URL'):
        value: Any = information.get(field)

        if isinstance(value, str):
            text_values.append(value)

    reference_data: Any = information.get('Reference', [])
    reference_values: list[Any]

    if isinstance(reference_data, str):
        reference_values = [reference_data]
    elif isinstance(reference_data, list):
        reference_values = reference_data
    else:
        reference_values = []

    for reference in reference_values:
        if isinstance(reference, str):
            text_values.append(reference)

    years: list[int] = []

    for text in text_values:
        matches: list[str] = re.findall(
            r'(?<!\d)(?:18|19|20|21)\d{2}(?!\d)',
            text
        )

        years.extend(int(match) for match in matches)

    return max(years) if years else None

def find_property(sections: list[dict[str, Any]], heading: str) -> dict[str, Any] | None: # AI Function

    property_records: list[dict[str, Any]] = collect_property_records(
        sections, heading
    )

    if not property_records:
        return None

    if heading in ('Melting Point', 'Boiling Point'):

        def pressure_rank(record: dict[str, Any]) -> int:

            value = record.get('value', {})
            entries = value.get('StringWithMarkup', []) if isinstance(value, dict) else []

            text = ' '.join(
                entry.get('String', '')
                for entry in entries
                if isinstance(entry.get('String'), str)
            )

            standard_pressure = re.search(
                r'\b(?:760\s*mm\s*Hg|1\s*atm|101325\s*Pa|101\.325\s*kPa|0\.101325\s*MPa)\b',
                text,
                re.IGNORECASE
            )

            any_pressure = re.search(
                r'\b(?:mm\s*Hg|atm|kPa|MPa|Pa|bar)\b',
                text,
                re.IGNORECASE
            )

            if standard_pressure:
                return 2

            if any_pressure:
                return 0

            return 1

        highest_rank = max(
            pressure_rank(record) for record in property_records
        )

        property_records = [
            record for record in property_records
            if pressure_rank(record) == highest_rank
        ]

    latest_record: dict[str, Any] | None = None

    for record in property_records:

        record_year: int | None = record['year']

        if record_year is None:
            continue

        if latest_record is None or record_year > latest_record['year']:
            latest_record = record

    return latest_record if latest_record is not None else property_records[0]
def fetch_chemical_records() -> list[tuple[str, str, dict[str, Any]]]:

    records: list[tuple[str, str, dict[str, Any]]] = []

    for chemical_input in chemicals:

        chemical, phase = identify_phase(chemical_input)

        try:

            if re.fullmatch(formula_pattern, chemical):
                chemical_data = chemical_formula_fetch(chemical, phase)

            else:
                chemical_data = chemical_name_fetch(chemical, phase)

        except requests.exceptions.RequestException as error:
            print(f'Unable to fetch {chemical}: {error}')
            continue

        if chemical_data is None:
            continue

        records.append((chemical, phase, chemical_data))

    return records
#Word Related Functions #AI
def format_table_value(value: Any, field: str, phase: str | None = None) -> str:
    if value is None:
        return 'Not determined'

    if field == 'danger' and isinstance(value, list):
        return ', '.join(value) if value else 'Not determined'
    if field == 'physical_description' and isinstance(value, list):

        color_pattern = re.compile(
            r'\b(colorless|colourless|white|yellow|red|orange|brown|black|blue|green|pink|purple|violet|gray|grey)\b',
            re.IGNORECASE
        )

        state_pattern = re.compile(
            r'\b(solid|liquid|gas|crystalline|crystals?|powder|vapor|vapour)\b',
            re.IGNORECASE
        )

        odor_pattern = re.compile(
            r'\b(?:odorless|odourless|'
            r'(?:[\w-]+-like|pungent|irritating|sharp|strong|sweet|fruity|fishy|musty)'
            r'(?:[\s,]+(?:pungent|irritating|sharp|strong|sweet|fruity|fishy|musty))*'
            r'\s+(?:odor|odour)|'
            r'(?:odor|odour)\s+(?:of|like)\s+[^,.;]+)',
            re.IGNORECASE
        )

        phase_patterns = {
            'gas': re.compile(r'\b(gas|vapou?r)\b', re.IGNORECASE),
            'liquid': re.compile(r'\bliquid\b', re.IGNORECASE),
            'solid': re.compile(r'\b(solid|crystals?|crystalline|powder)\b', re.IGNORECASE)
        }

        phase_pattern = phase_patterns.get((phase or '').lower())

        best_score = (-1, -1)
        best_description = []

        for record in value:

            if not isinstance(record, dict):
                continue

            text = record.get('value', '')

            if not isinstance(text, str) or not text.strip():
                continue

            color = color_pattern.search(text)
            state = state_pattern.search(text)
            odor = odor_pattern.search(text)

            parts = []

            if color:
                color_name = color.group().lower()

                if color_name == 'colourless':
                    color_name = 'colorless'

                parts.append(color_name.capitalize())

            if state:
                parts.append(state.group().lower())

            if odor:
                parts.append(odor.group().lower())

            if not parts:
                continue

            phase_match = int(
                bool(phase_pattern and phase_pattern.search(text))
            )

            score = (phase_match, len(parts))

            if score > best_score:
                best_score = score
                best_description = parts

        return ', '.join(best_description) or 'Not determined'

    if field == 'first_aid_measures' and isinstance(value, list):
        preferred_records = [
            record for record in value
            if 'International Chemical Safety Cards' in (
                record.get('source_name') or ''
            )
        ]

        if not preferred_records:
            return 'Not determined'

        preferred_order = {
            'Inhalation First Aid': 1,
            'Skin First Aid': 2,
            'Eye First Aid': 3,
            'Ingestion First Aid': 4
        }

        preferred_records.sort(
            key=lambda record: preferred_order.get(record.get('name'), 99)
        )

        return '\n\n'.join(
            f"{record['name'].replace(' First Aid', '')}: {record['value']}"
            for record in preferred_records
        )

    if field in ('melting_point', 'boiling_point') and isinstance(value, str):
        value = re.sub(
            r'\s*at\s*(?:760\s*mm\s*Hg|1\s*atm|101325\s*Pa|101\.325\s*kPa|0\.101325\s*MPa)\b',
            '',
            value,
            flags=re.IGNORECASE
        )
        return value.strip()
    if isinstance(value, str):
        return value

    if isinstance(value, dict):

        property_value: Any = value.get('value')

        if isinstance(property_value, dict):

            string_parts: list[dict[str, Any]] = (
                property_value.get('StringWithMarkup', [])
            )

            values: list[str] = [
                part['String']
                for part in string_parts
                if isinstance(part.get('String'), str)
            ]

            if field == 'density':
                cleaned_values: list[str] = []

                for item in values:
                    match = re.search(
                        r'([<>≤≥~]?\s*\d+(?:\.\d+)?(?:\s*[-–]\s*\d+(?:\.\d+)?)?)\s*(?:g/cu cm|g/cm³)',
                        item,
                        re.IGNORECASE
                    )

                    if match:
                        cleaned_values.append(f'{match.group(1)} g/cm³')

                return '; '.join(cleaned_values) if cleaned_values else 'Not determined'

            return '; '.join(values) if values else 'Not determined'

        if isinstance(property_value, str):
            return property_value

    return str(value)
def add_ghs_pictograms(
    cell: Any,
    pictogram_codes: list[str] | str | None,
    image_cache: dict[str, bytes]
) -> None: # AI Function

    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not isinstance(pictogram_codes, list):
        paragraph.add_run(str(pictogram_codes or 'Not determined'))
        return

    if not pictogram_codes:
        paragraph.add_run('Not determined')
        return

    for code in pictogram_codes:

        if code not in image_cache:

            image_url: str = (
                f'https://pubchem.ncbi.nlm.nih.gov/images/ghs/{code}.gif'
            )

            try:
                response = requests.get(image_url, timeout=10)
                response.raise_for_status()

                image_cache[code] = response.content

            except requests.exceptions.RequestException as error:
                print(f'Unable to retrieve {code}: {error}')
                paragraph.add_run(code)
                continue

        run = paragraph.add_run()

        run.add_picture(
            BytesIO(image_cache[code]),
            width=Inches(0.25)
        )
def add_first_aid_measures(cell: Any, value: Any) -> None:
    cell.text = ''

    preferred_records = [
        record for record in value
        if isinstance(record, dict)
        and 'International Chemical Safety Cards' in (
            record.get('source_name') or ''
        )
        and record.get('value')
    ] if isinstance(value, list) else []

    preferred_order = {
        'Inhalation First Aid': 1,
        'Skin First Aid': 2,
        'Eye First Aid': 3,
        'Ingestion First Aid': 4
    }

    preferred_records.sort(
        key=lambda record: preferred_order.get(record.get('name'), 99)
    )

    if not preferred_records:
        run = cell.paragraphs[0].add_run('Not determined')
        run.font.size = Pt(7)
        return

    for index, record in enumerate(preferred_records):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()

        paragraph.paragraph_format.space_after = Pt(2)

        label = record['name'].replace(' First Aid', '')
        instructions = record['value'].strip()

        label_run = paragraph.add_run(f'{label}: ')
        label_run.bold = True
        label_run.font.size = Pt(7)

        instruction_run = paragraph.add_run(instructions)
        instruction_run.font.size = Pt(7)
def select_sds_content() -> list[str]: # AI Function

    options: dict[int, tuple[str, str]] = {
        1: ('Molecular Weight', 'molecular_weight'),
        2: ('Melting Point', 'melting_point'),
        3: ('Boiling Point', 'boiling_point'),
        4: ('Density', 'density'),
        5: ('GHS Pictograms', 'danger'),
        6: ('Physical Description', 'physical_description'),
        7: ('First-Aid Measures', 'first_aid_measures'),
    }

    selected_content: list[str] = []

    print('\nSelect SDS content in your preferred order.')

    while len(selected_content) < len(options):

        print('\nAvailable content:')

        for number, (label, key) in options.items():

            if key not in selected_content:
                print(f'{number}. {label}')

        choice: str = input(
            f'\nWhat should item #{len(selected_content) + 1} be? '
            '(Press Enter to finish): '
        ).strip()

        if not choice:
            break

        if not choice.isdigit() or int(choice) not in options:
            print('Invalid selection. Try again.')
            continue

        selected_key: str = options[int(choice)][1]

        if selected_key in selected_content:
            print('Already selected. Choose another item.')
            continue

        selected_content.append(selected_key)

    return selected_content

#Main fetching functions #AI
# Main fetching functions

def chemical_data_fetch(cid: int, chemical: str, phase: str) -> dict[str, Any]:
    # Molecular Weight
    molecular_weight_url = (
        f'{pug_rest_base_url}/compound/cid/'
        f'{cid}/property/MolecularWeight/JSON'
    )

    molecular_weight_data = fetch_optional_data(molecular_weight_url)

    molecular_weight_properties = (
        molecular_weight_data.get('PropertyTable', {}).get('Properties', [])
        if molecular_weight_data else []
    )

    molecular_weight = (
        molecular_weight_properties[0].get('MolecularWeight')
        if molecular_weight_properties else None
    )

    # Experimental Properties
    experimental_url = (
        f'{pug_view_base_url}/data/compound/'
        f'{cid}/JSON?heading=Experimental+Properties'
    )

    experimental_data = fetch_optional_data(experimental_url)

    experimental_record = (
        experimental_data.get('Record', {})
        if experimental_data else {}
    )

    experimental_sections = experimental_record.get('Section', [])
    experimental_references = experimental_record.get('Reference', [])

    physical_description = collect_information_records(
        experimental_sections,
        'Physical Description',
        experimental_references
    )

    melting_point = find_property(
        experimental_sections, 'Melting Point'
    )

    boiling_point = find_property(
        experimental_sections, 'Boiling Point'
    )

    density = find_property(
        experimental_sections, 'Density'
    )

    # Safety and Hazards
    safety_url = (
        f'{pug_view_base_url}/data/compound/'
        f'{cid}/JSON?heading=Safety+and+Hazards'
    )

    safety_data = fetch_optional_data(safety_url)

    safety_record = (
        safety_data.get('Record', {})
        if safety_data else {}
    )

    safety_sections = safety_record.get('Section', [])
    safety_references = safety_record.get('Reference', [])

    danger = (
        find_ghs_symbols(safety_sections)
        if safety_data else None
    )

    if danger and phase != 'gas':
        danger = [code for code in danger if code != 'GHS04']

    first_aid_measures = collect_information_records(
        safety_sections,
        'First Aid Measures',
        safety_references
    )

    return {
        'chemical': chemical,
        'cid': cid,
        'molecular_weight': molecular_weight if molecular_weight is not None else 'Not determined',
        'melting_point': melting_point if melting_point is not None else 'Not determined',
        'boiling_point': boiling_point if boiling_point is not None else 'Not determined',
        'density': density if density is not None else 'Not determined',
        'danger': danger if danger is not None else 'Not determined',
        'physical_description': physical_description,
        'first_aid_measures': first_aid_measures,
    }
def chemical_formula_fetch(chemical_formula: str, phase: str) -> dict[str, Any] | None:

    encoded_formula = urllib.parse.quote(chemical_formula, safe='')

    formula_url = (
        f'{pug_rest_base_url}/compound/fastformula/'
        f'{encoded_formula}/cids/JSON?MaxRecords=100'
    )

    chemical_data = check_chemical_if_found(formula_url)

    if chemical_data is None:
        chemical_not_found(chemical_formula)
        return None

    cid_list = chemical_data.get('IdentifierList', {}).get('CID', [])

    if not cid_list:
        chemical_not_found(chemical_formula)
        return None

    cid = cid_list[0]

    return chemical_data_fetch(cid, chemical_formula, phase)
def chemical_name_fetch(chemical_name: str, phase: str) -> dict[str, Any] | None:

    encoded_name = urllib.parse.quote(chemical_name, safe='')

    name_url = (
        f'{pug_rest_base_url}/compound/name/'
        f'{encoded_name}/cids/JSON'
    )

    chemical_data = check_chemical_if_found(name_url)

    if chemical_data is None:
        chemical_not_found(chemical_name)
        return None

    cid_list = chemical_data.get('IdentifierList', {}).get('CID', [])

    if not cid_list:
        chemical_not_found(chemical_name)
        return None

    if len(cid_list) > 1:
        print(
            f'Multiple PubChem matches found for {chemical_name}. '
            f'Using the first result (CID {cid_list[0]}).'
        )

    cid = cid_list[0]

    return chemical_data_fetch(cid, chemical_name, phase)

chemicals: list[str] = [] #User Input
while True:
    chemical: str = input('Enter chemical (press Enter to finish): ').strip()
    if not chemical:
        break
    chemicals.append(chemical)

if not check_server(): #check if pubchem is online
    input("Press Enter to exit.")
    raise SystemExit

# # My Output
# # for chemical in chemicals:
# #     chemical, phase = identify_phase(chemical)
# #     if re.fullmatch(formula_pattern, chemical): 
# #         chemical_formula_fetch(chemical, phase) #Calls function for requesting at pubchem with chemical formula
# #     else:
# #         chemical_name_fetch(chemical, phase) #Calls function for requesting at pubchem with chemical name

# Output
# Fetch chemical data in the background

# Run PubChem requests while prompting for SDS content
with ThreadPoolExecutor(max_workers=1) as executor:

    fetch_future = executor.submit(fetch_chemical_records)

    selected_content: list[str] = select_sds_content()

    chemical_records = fetch_future.result()


table_column_keys: set[str] = {
    'molecular_weight',
    'melting_point',
    'boiling_point',
    'density',
    'danger',
    'physical_description',
    'first_aid_measures'
}


selected_columns: list[str] = [
    key for key in selected_content
    if key in table_column_keys
]



# Word Processing
document = Document()

style = document.styles['Normal']
style.font.name = 'Aptos'
style.font.size = Pt(10.5)

column_labels: dict[str, str] = {
    'molecular_weight': 'Molecular Weight',
    'melting_point': 'Melting Point (°C)',
    'boiling_point': 'Boiling Point (°C)',
    'density': 'Density',
    'danger': 'GHS Pictograms',
    'physical_description': 'Physical Description',
    'first_aid_measures': 'First-Aid Measures'
}

table_fields: list[tuple[str, str]] = [
    ('chemical', 'Compound')
] + [
    (key, column_labels[key])
    for key in selected_columns
]

table = document.add_table(
    rows=1,
    cols=len(table_fields)
)

table.style = 'Table Grid'

for cell, (field, heading) in zip(table.rows[0].cells, table_fields):

    paragraph = cell.paragraphs[0]

    if field == 'molecular_weight':
        run = paragraph.add_run('M')
        run.bold = True

        run = paragraph.add_run('W')
        run.bold = True
        run.font.subscript = True

        run = paragraph.add_run(' (g/mol)')
        run.bold = True

    else:
        run = paragraph.add_run(heading)
        run.bold = True

section_cells: list[Any] = []
# Populate table
for chemical, phase, chemical_data in chemical_records:

    row = table.add_row()

    display_name: str = (
        chemical if phase == 'aqueous'
        else f'{chemical} ({phase})'
    )

    compound_paragraph = row.cells[0].paragraphs[0]

    for part in re.split(r'(\d+)', display_name):
        run = compound_paragraph.add_run(part)

        if part.isdigit():
            run.font.subscript = True

    # Cache downloaded GHS images
    image_cache: dict[str, bytes] = {}

    for index, field in enumerate(selected_columns, start=1):

        value: Any = chemical_data.get(field)

        if field == 'danger':
            add_ghs_pictograms(row.cells[index], value, image_cache)

        elif field == 'first_aid_measures':
            add_first_aid_measures(row.cells[index], value)

        else:
            row.cells[index].text = format_table_value(value, field, phase)
 
# Center table contents
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Left-align section headings and descriptions
for cell in section_cells:
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

# Save Word document
root: Tk = Tk()
root.withdraw()

save_path: str = filedialog.asksaveasfilename(
    defaultextension='.docx',
    filetypes=[('Word Document', '*.docx')],
    title='Where to save SDS Document?'
)

if save_path:
    document.save(save_path)

root.destroy()

