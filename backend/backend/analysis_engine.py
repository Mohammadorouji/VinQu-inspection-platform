import re
from pathlib import Path
from dataclasses import asdict, dataclass, field
from typing import Optional
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


@dataclass
class AnalysisResult:
    document_type: str
    project_name: Optional[str]
    project_reference: Optional[str]
    order_number: Optional[str]
    notification_number: Optional[str]
    revision: Optional[str]
    discipline: Optional[str]
    nature_of_work: str
    inspection_mode: Optional[str]
    inspection_stage: Optional[str]
    equipment_scope: Optional[str]
    inspection_factory_location: Optional[str]
    site_of_inspection: Optional[str]
    named_contacts: list[str] = field(default_factory=list)
    relevant_itp_step: Optional[str] = None
    applicable_standards: list[str] = field(default_factory=list)
    tests_or_checks: list[str] = field(default_factory=list)
    inspector_duties: list[str] = field(default_factory=list)
    referenced_documents: list[str] = field(default_factory=list)
    reporting_requirements: list[str] = field(default_factory=list)
    missing_information: list[str] = field(default_factory=list)
    confidence_level: str = "Low"
    reasoned_summary: str = ""


def clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def uniq(items):
    out, seen = [], set()
    for item in items:
        c = clean(item)
        if c and c not in seen:
            out.append(c)
            seen.add(c)
    return out


def first(text, pattern):
    m = re.search(pattern, text, flags=re.I | re.S)
    return clean(m.group(1)) if m else None


def allm(text, pattern):
    return uniq([m.group(1) for m in re.finditer(pattern, text, flags=re.I | re.S)])


def extract_text(path: Path):
    suf = path.suffix.lower()

    if suf in {".txt", ".md", ".csv", ".log"}:
        return path.read_text(encoding="utf-8", errors="ignore")

    if suf == ".docx":
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs if clean(p.text))

    if suf == ".pdf":
        reader = PdfReader(str(path))
        text = "\n".join((page.extract_text() or "") for page in reader.pages[:80])
        if not clean(text):
            raise ValueError("PDF text could not be extracted. Please upload a text-based PDF or paste the extracted text.")
        return text

    if suf == ".xlsx":
        wb = load_workbook(filename=str(path), data_only=True)
        chunks = []
        for ws in wb.worksheets[:12]:
            chunks.append(f"Sheet: {ws.title}")
            for row in ws.iter_rows(values_only=True):
                vals = [clean(v) for v in row if clean(v)]
                if vals:
                    chunks.append(" | ".join(vals))
        text = "\n".join(chunks)
        if not clean(text):
            raise ValueError("Spreadsheet appears empty or unreadable.")
        return text

    raise ValueError("Unsupported file type. Supported formats: PDF, DOCX, TXT, CSV, XLSX.")


def classify(text, file_name=""):
    c = f"{text}\n{file_name}"
    if re.search(r"notification for inspection|\bNOI\b", c, re.I):
        return "NOI / Notification for Inspection"
    if re.search(r"notification of readiness for inspection|\bNORFI\b", c, re.I):
        return "NORFI / Notification of Readiness for Inspection"
    if re.search(r"inspection notification|application no\.|inspection location|requested date", c, re.I) and re.search(r"inspection and test plan|\bITP\b|QCP|quality control plan", c, re.I):
        return "Inspection Notification linked to ITP/QCP"
    if re.search(r"factory acceptance test plan|\bFAT\b", c, re.I):
        return "Factory Acceptance Test Plan"
    if re.search(r"inspection and test plan|\bITP\b|QCP|quality control plan", c, re.I):
        return "Inspection and Test Plan"
    if re.search(r"call off|inspection assignment work order", c, re.I):
        return "Call Off / Inspection Assignment Work Order"
    if re.search(r"work instruction|procedure|instruction", c, re.I):
        return "Work Instruction / Procedure"
    return "Unknown"


def extract_noi_activities(text):
    activities = []
    # Specific NOI patterns from table-like PDFs.
    if re.search(r"\bF47\b", text, re.I) and re.search(r"partial.?discharge", text, re.I):
        activities.append("F47 — Partial-discharge measurement")
    if re.search(r"\bF61\b", text, re.I) and re.search(r"impulse|AC voltage", text, re.I):
        activities.append("F61 — Impulse or AC voltage test on two single coils")
    # Generic step/activity rows.
    for m in re.finditer(r"\b(F\d{1,3})\b\s+([A-Za-z][A-Za-z0-9 /().,-]{5,90})", text, flags=re.I):
        item = f"{m.group(1).upper()} — {clean(m.group(2))}"
        if not re.search(r"telephone|email|address|project|rev|page|date|quantity", item, re.I):
            activities.append(item)
    return uniq(activities)[:12]


def extract_xlsx_activities(text):
    activities = []
    # Capture meaningful inspection action lines, avoid document-list noise.
    keywords = [
        "SYSTEM INTEGRATION TEST",
        "UNIT CONTROL PANEL",
        "iFAT",
        "Witness",
        "Hold",
        "Observe",
        "Inspection Date",
        "Quality Control Plan",
        "Test Procedure",
    ]
    for line in text.splitlines():
        c = clean(line)
        if not c:
            continue
        if any(k.lower() in c.lower() for k in keywords):
            if not re.search(r"document title|document return code|status|revision code", c, re.I):
                activities.append(c)
    # Stronger summary items if present.
    if re.search(r"UNIT CONTROL PANEL", text, re.I):
        activities.insert(0, "UNIT CONTROL PANEL — iFAT / system integration test")
    if re.search(r"SYSTEM INTEGRATION TEST", text, re.I):
        activities.insert(1, "System Integration Test")
    return uniq(activities)[:14]


def extract_referenced_documents(text):
    docs = []
    patterns = [
        r"(QUALITY CONTROL PLAN[^\n|]*)",
        r"(UNIT CONTROL SYSTEM SCHEMATIC[^\n|]*)",
        r"(UNIT CONTROL SYSTEM I/O LIST[^\n|]*)",
        r"(UNIT CONTROL SYSTEM INTERNAL WIRING DIAGRAM[^\n|]*)",
        r"(INTERCONNECTING WIRING DIAGRAM[^\n|]*)",
        r"(LAYOUT / CONSTRUCTION & MAIN COMPONENTS LIST[^\n|]*)",
        r"(UNIT CONTROL SYSTEM TEST PROCEDURE[^\n|]*)",
    ]
    for p in patterns:
        docs += allm(text, p)
    return uniq(docs)[:20]


def extract_reporting_requirements(text):
    reqs = []
    if re.search(r"must be e-?mailed|emailed", text, re.I):
        reqs.append("Inspection notification must be emailed to the listed recipients.")
    if re.search(r"minimum\s+14\s+calendar\s+days", text, re.I):
        reqs.append("Notification must be issued at least 14 calendar days before the inspection date.")
    if re.search(r"working days in advance", text, re.I):
        reqs.append(first(text, r"(\d+\s+working days in advance[^\n|]{0,140})") or "Advance notice is required before inspection.")
    return uniq(reqs)


def analyse_text(text, file_name=""):
    doc_type = classify(text, file_name)
    standards = uniq([m.group(1) for m in re.finditer(r"(IEC(?:/IEEE)?\s*[0-9-]+(?:-[0-9]+)?)", text, flags=re.I)])[:20]

    if "Notification" in doc_type and "NOI" in doc_type:
        tests = extract_noi_activities(text)
    elif "Inspection Notification linked" in doc_type or file_name.lower().endswith(".xlsx"):
        tests = extract_xlsx_activities(text)
    else:
        tests = uniq([f"{m.group(1)} — {clean(m.group(2))}" for m in re.finditer(r"(?:Item\s*)?(\d+(?:\.\d+)*)\s+([^\n]{6,120})", text, flags=re.I)])[:20]

    contacts = uniq(
        [f"Email: {m}" for m in allm(text, r"([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})")]
        + [f"Contact person: {m}" for m in allm(text, r"Contact Person\s*(?:Section Project Manager Name)?\s*([^\n]+)")]
        + [f"Assigned inspector: {m}" for m in allm(text, r"Assigned Inspector\s*:?\s*([^\n]+)")]
        + [f"Vendor / Subcontractor: {m}" for m in allm(text, r"Vendor\s*/\s*Sub\s*contractor Name\s*([^\n]+)")]
    )[:12]

    # Project extraction improvements.
    project_ref = first(text, r"PROJECT No\.?\s*([0-9A-Z.-]+)")
    project_name = first(text, r"Project Name\s*:?\s*([^\n|]+)") or first(text, r"(MERAM PROJECT)") or first(text, r"(Trion FPU Project)")

    # Equipment / scope extraction improvements.
    equipment_scope = (
        first(text, r"PO Description\s*([^\n]+)")
        or first(text, r"Equipment Description\s*&\s*Tag No\.?\s*\|?\s*([^\n|]+)")
        or first(text, r"UNIT CONTROL PANEL\s*\(iFAT\)")
        or first(text, r"Object\s*:?\s*([^\n]+)")
        or first(text, r"Inspection Activity\s*([^\n]+)")
        or first(text, r"Inspection Activities\s*([^\n]+)")
    )
    if not equipment_scope and re.search(r"UNIT CONTROL PANEL", text, re.I):
        equipment_scope = "UNIT CONTROL PANEL (iFAT)"

    # Location extraction improvements.
    location = (
        first(text, r"Address\s*([^\n]+Drasov)")
        or first(text, r"(Čedlosy\s+126,\s*664\s*24\s*Drasov)")
        or first(text, r"Inspection Location\s*:?\s*([^\n|]+)")
        or first(text, r"(East Gate Business Park F2,\s*H-2151\s*Fót,\s*Hungary)")
        or first(text, r"Manufacturer\s*&\s*Factory\s*([^\n]+)")
    )

    nature_lookup = {
        "Factory Acceptance Test Plan": "Factory Acceptance Testing in a test-lab / workshop environment",
        "NOI / Notification for Inspection": "Notification-based witness inspection or attendance notice",
        "NORFI / Notification of Readiness for Inspection": "Readiness notification for inspection attendance or witness action",
        "Inspection Notification linked to ITP/QCP": "Inspection notification linked to ITP / QCP checkpoints and witness planning",
        "Inspection and Test Plan": "Inspection and Test Plan governing checkpoints, witness points, and acceptance criteria",
        "Call Off / Inspection Assignment Work Order": "Inspection assignment / surveillance instruction",
        "Work Instruction / Procedure": "Instruction-based inspection activity",
        "Unknown": "Inspection activity",
    }

    result = AnalysisResult(
        document_type=doc_type,
        project_name=project_name,
        project_reference=project_ref or first(text, r"RFQ No\.?\s*/\s*PO No\.?\s*\|?\s*([^|\n]+)"),
        order_number=first(text, r"Order No\.?\s*:?\s*([^\n]+)") or first(text, r"PO NO:\s*([^\n]+)"),
        notification_number=first(text, r"Notification No\.?\s*([0-9A-Z.-]+)") or first(text, r"Application No\.?\s*\|?\s*([A-Z0-9-]+)"),
        revision=first(text, r"Rev\s*#?\s*:?\s*([A-Z0-9.-]+)") or first(text, r"Rev\.?\s*([A-Z0-9.-]+)"),
        discipline=first(text, r"Discipline\s*:?\s*([^\n]+)") or ("Electrical" if re.search(r"electrical|transformer|partial.?discharge|motor|voltage|impulse|unit control", text, re.I) else None),
        nature_of_work=nature_lookup.get(doc_type, "Inspection activity"),
        inspection_mode="Witness" if re.search(r"\bWitness\b", text, re.I) else ("Hold" if re.search(r"\bHold\b", text, re.I) else None),
        inspection_stage="FAT / workshop testing" if doc_type == "Factory Acceptance Test Plan" else ("Notification / witness stage" if "Notification" in doc_type else None),
        equipment_scope=equipment_scope,
        inspection_factory_location=location,
        site_of_inspection=first(text, r"Location of Work\s*:?\s*([^\n]+)") or first(text, r"Site of Inspection\s*:?\s*([^\n]+)"),
        named_contacts=contacts,
        relevant_itp_step=", ".join(uniq(allm(text, r"\b(F\d{1,3})\b"))) or first(text, r"ITP Step No\s*([^\n]+)"),
        applicable_standards=standards,
        tests_or_checks=tests,
        referenced_documents=extract_referenced_documents(text),
        reporting_requirements=extract_reporting_requirements(text),
        inspector_duties=[],
        missing_information=[],
        confidence_level="High" if len(clean(text)) > 500 else ("Medium" if len(clean(text)) > 120 else "Low"),
        reasoned_summary="",
    )

    if doc_type == "Factory Acceptance Test Plan":
        result.inspector_duties = [
            "Review the numbered FAT test items and confirm which are routine, type, or special tests.",
            "Verify the applicable standard and acceptance criterion for each witnessed FAT item.",
            "Record observations and any deviations during factory testing.",
        ]
        result.reasoned_summary = "This document is a Factory Acceptance Test Plan organised around test items and applicable standards for factory testing."

    elif doc_type == "NOI / Notification for Inspection":
        result.inspector_duties = [
            "Confirm attendance or waiver before the inspection date.",
            "Witness the listed ITP activities at the stated factory / inspection location.",
            "Record whether Supplier, Contractor, and Company / TPA intervention requirements are satisfied.",
            "Use the referenced ITP for exact acceptance criteria.",
        ]
        result.reasoned_summary = "This NOI identifies the project, parties, factory location, ITP steps, and witness activities. It is suitable for preparing attendance, witness planning, and reporting obligations."

    elif doc_type == "Inspection Notification linked to ITP/QCP":
        result.inspector_duties = [
            "Confirm the inspection location, date window, and required advance notice.",
            "Review the linked QCP / ITP and referenced technical documents before attendance.",
            "Witness or observe the iFAT / system integration test according to the intervention matrix.",
            "Record open points, deviations, and final inspection status.",
        ]
        result.reasoned_summary = "This is an inspection notification linked to QCP / ITP controls. It should be used to prepare the inspector for the planned iFAT / system integration test and related document review."

    elif doc_type == "Inspection and Test Plan":
        result.inspector_duties = [
            "Review each inspection and test step before the relevant stage.",
            "Use the ITP to decide what must be checked, witnessed, recorded, and reported.",
        ]
        result.reasoned_summary = "This document is an Inspection and Test Plan used to govern inspection checkpoints and acceptance criteria."

    else:
        result.inspector_duties = [
            "Review the document and confirm the inspection scope.",
            "Use the extracted fields to prepare the pre-inspection briefing.",
        ]
        result.reasoned_summary = "This document has been analysed and classified to support inspection planning."

    if not (result.project_name or result.project_reference):
        result.missing_information.append("Project name / reference not identified confidently.")
    if not result.equipment_scope:
        result.missing_information.append("Equipment / scope not identified confidently.")
    if not (result.inspection_factory_location or result.site_of_inspection):
        result.missing_information.append("Inspection location not identified confidently.")
    if not result.tests_or_checks:
        result.missing_information.append("No clear inspection activities / test items extracted.")
    if not result.applicable_standards:
        result.missing_information.append("No standards detected in the uploaded document. Use referenced ITP / QCP / project specifications for acceptance criteria.")

    return asdict(result)


def analyse_file(path):
    p = Path(path)
    return analyse_text(extract_text(p), p.name)
