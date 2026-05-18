import re
from pathlib import Path
from dataclasses import asdict, dataclass, field
from typing import Optional
from pypdf import PdfReader
from docx import Document

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
    missing_information: list[str] = field(default_factory=list)
    confidence_level: str = "Low"
    reasoned_summary: str = ""

def clean(text): return re.sub(r"\s+", " ", text or "").strip()
def uniq(items):
    out=[]; seen=set()
    for item in items:
        c=clean(item)
        if c and c not in seen:
            out.append(c); seen.add(c)
    return out
def first(text, pattern):
    m=re.search(pattern, text, flags=re.I|re.S)
    return clean(m.group(1)) if m else None
def allm(text, pattern):
    return uniq([m.group(1) for m in re.finditer(pattern, text, flags=re.I|re.S)])

def extract_text(path: Path):
    suf=path.suffix.lower()
    if suf in {".txt",".md",".csv",".log"}:
        return path.read_text(encoding="utf-8", errors="ignore")
    if suf==".docx":
        doc=Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs if clean(p.text))
    if suf==".pdf":
        reader=PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages[:50])
    raise ValueError("Unsupported file type")

def classify(text, file_name=""):
    c=f"{text}\n{file_name}"
    if re.search(r"factory acceptance test plan|\bFAT\b", c, re.I): return "Factory Acceptance Test Plan"
    if re.search(r"notification of readiness for inspection|\bNORFI\b", c, re.I): return "NORFI / Notification of Readiness for Inspection"
    if re.search(r"notification for inspection", c, re.I): return "NOI / Notification for Inspection"
    if re.search(r"inspection and test plan|\bITP\b", c, re.I): return "Inspection and Test Plan"
    if re.search(r"call off|inspection assignment work order", c, re.I): return "Call Off / Inspection Assignment Work Order"
    if re.search(r"work instruction|procedure|instruction", c, re.I): return "Work Instruction / Procedure"
    return "Unknown"

def analyse_text(text, file_name=""):
    doc_type=classify(text, file_name)
    standards=uniq([m.group(1) for m in re.finditer(r"(IEC(?:/IEEE)?\s*[0-9-]+(?:-[0-9]+)?)", text, flags=re.I)])[:20]
    tests=uniq([f"{m.group(1)} — {clean(m.group(2))}" for m in re.finditer(r"(?:Item\s*)?(\d+(?:\.\d+)*)\s+([^\n]{6,120})", text, flags=re.I)])[:20]
    contacts=uniq([f"Email: {m}" for m in allm(text, r"([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})")] +
                  [f"Contact person: {m}" for m in allm(text, r"Contact Person\s*:?\s*([^\n]+)")] +
                  [f"Assigned inspector: {m}" for m in allm(text, r"Assigned Inspector\s*:?\s*([^\n]+)")])[:12]
    result=AnalysisResult(
        document_type=doc_type,
        project_name=first(text, r"Project Name\s*:?\s*([^\n]+)"),
        project_reference=first(text, r"PROJECT No\.\s*([^\n]+)") or first(text, r"Customer\s*/\s*Project reference\s*:?\s*([^\n]+)"),
        order_number=first(text, r"Order No\.?\s*:?\s*([^\n]+)"),
        notification_number=first(text, r"Notification No\.?\s*([^\n]+)"),
        revision=first(text, r"Rev\.?\s*([A-Z0-9.-]+)"),
        discipline=first(text, r"Discipline\s*:?\s*([^\n]+)") or ("Electrical" if re.search(r"electrical|transformer|partial.?discharge|motor", text, re.I) else None),
        nature_of_work={{
            "Factory Acceptance Test Plan":"Factory Acceptance Testing in a test-lab / workshop environment",
            "NOI / Notification for Inspection":"Notification-based witness inspection or attendance notice",
            "NORFI / Notification of Readiness for Inspection":"Readiness notification for inspection attendance or witness action",
            "Inspection and Test Plan":"Inspection and Test Plan governing checkpoints, witness points, and acceptance criteria",
            "Call Off / Inspection Assignment Work Order":"Inspection assignment / surveillance instruction",
            "Work Instruction / Procedure":"Instruction-based inspection activity",
            "Unknown":"Inspection activity"
        }}.get(doc_type, "Inspection activity"),
        inspection_mode="Witness" if re.search(r"\bWitness\b", text, re.I) else None,
        inspection_stage="FAT / workshop testing" if doc_type=="Factory Acceptance Test Plan" else ("Notification / witness stage" if "Notification" in doc_type else None),
        equipment_scope=first(text, r"Object\s*:?\s*([^\n]+)") or first(text, r"Inspection Activity\s*([^\n]+)") or first(text, r"Item description\s*:?\s*([^\n]+)"),
        inspection_factory_location=first(text, r"Place of Witness[\s\S]*?Address\s*([^\n]+)") or first(text, r"Address\s*:?\s*([^\n]+)"),
        site_of_inspection=first(text, r"Location of Work\s*:?\s*([^\n]+)") or first(text, r"Site of Inspection\s*:?\s*([^\n]+)"),
        named_contacts=contacts,
        relevant_itp_step=first(text, r"ITP Step No\s*([^\n]+)"),
        applicable_standards=standards,
        tests_or_checks=tests,
        inspector_duties=[],
        missing_information=[],
        confidence_level="High" if len(clean(text))>500 else ("Medium" if len(clean(text))>120 else "Low"),
        reasoned_summary=""
    )
    if doc_type=="Factory Acceptance Test Plan":
        result.inspector_duties=[
            "Review the numbered FAT test items and confirm which are routine, type, or special tests.",
            "Verify the applicable standard and acceptance criterion for each witnessed FAT item.",
            "Record observations and any deviations during factory testing."
        ]
        result.reasoned_summary="This document is a Factory Acceptance Test Plan organised around test items and applicable standards for factory testing."
    elif "Notification" in doc_type:
        result.inspector_duties=[
            "Confirm attendance or waiver before the inspection date.",
            "Identify the site, parties, date, witness activity, and linked ITP step.",
            "Attend the inspection and record findings against the referenced activity."
        ]
        result.reasoned_summary="This document is a notification-type inspection document. Its role is to identify the site, parties, date, and witness activity before attendance."
    else:
        result.inspector_duties=[
            "Review the document and confirm the inspection scope.",
            "Use the extracted fields to prepare the pre-inspection briefing."
        ]
        result.reasoned_summary="This document has been analysed and classified to support inspection planning."
    if not (result.project_name or result.project_reference): result.missing_information.append("Project name / reference not identified confidently.")
    if not result.equipment_scope: result.missing_information.append("Equipment / scope not identified confidently.")
    if not (result.inspection_factory_location or result.site_of_inspection): result.missing_information.append("Inspection location not identified confidently.")
    if not result.tests_or_checks: result.missing_information.append("No clear test items extracted.")
    if not result.applicable_standards: result.missing_information.append("No standards detected.")
    return asdict(result)

def analyse_file(path):
    p=Path(path)
    return analyse_text(extract_text(p), p.name)
