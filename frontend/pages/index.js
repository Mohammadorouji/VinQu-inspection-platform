import { useMemo, useState } from "react";

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [selectedFile, setSelectedFile] = useState(null);
  const [referencedDocs, setReferencedDocs] = useState([]);
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Ready.");

  const [inspectionPhoto, setInspectionPhoto] = useState(null);
  const [photoPreview, setPhotoPreview] = useState("");
  const [photoAnalysis, setPhotoAnalysis] = useState(null);
  const [photoLoading, setPhotoLoading] = useState(false);
  const [inspectionArea, setInspectionArea] = useState("");
  const [inspectorNotes, setInspectorNotes] = useState("");
  const [findings, setFindings] = useState([]);
  const [currentItemIndex, setCurrentItemIndex] = useState(0);

  const inspectionItems = useMemo(() => result?.tests_or_checks || [], [result]);
  const currentItem = inspectionItems[currentItemIndex] || "General inspection item";

  function onInstructionFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setSelectedFile(file);
    setFileName(file.name);
    setResult(null);
    setError("");
    setStatus("Instruction file selected. Click Analyse document.");
  }

  function onReferencedDocs(e) {
    const files = Array.from(e.target.files || []);
    setReferencedDocs(files);
    setStatus(files.length ? `${files.length} referenced document(s) added.` : "No referenced documents added.");
  }

  async function analyseDocument() {
    setLoading(true);
    setError("");
    try {
      if (!selectedFile) throw new Error("Choose an instruction file first.");
      const form = new FormData();
      form.append("file", selectedFile);
      referencedDocs.forEach((file) => form.append("referenced_docs", file));
      setStatus("Uploading and analysing instruction file...");
      const res = await fetch(process.env.NEXT_PUBLIC_BACKEND_URL + "/analyse/file", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "File analysis failed.");
      setResult(data);
      setCurrentItemIndex(0);
      setStatus("Instruction analysed successfully.");
    } catch (err) {
      setError(err.message || "Analysis failed.");
      setStatus("Analysis failed.");
    } finally {
      setLoading(false);
    }
  }

  function onPhoto(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setInspectionPhoto(file);
    setPhotoPreview(URL.createObjectURL(file));
    setPhotoAnalysis(null);
    setStatus("Inspection photo selected. Click Analyse inspection photo.");
  }

  async function analysePhoto() {
    setPhotoLoading(true);
    setError("");
    try {
      if (!inspectionPhoto) throw new Error("Take or upload an inspection photo first.");
      const form = new FormData();
      form.append("photo", inspectionPhoto);
      form.append("inspection_area", inspectionArea);
      form.append("inspection_item", currentItem);
      form.append("inspector_notes", inspectorNotes);
      form.append("document_context", JSON.stringify(result || {}));
      const res = await fetch(process.env.NEXT_PUBLIC_BACKEND_URL + "/analyse/photo", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Photo analysis failed.");
      setPhotoAnalysis(data);
      if (!inspectorNotes) {
        setInspectorNotes("Photo received. Automated defect classification is not active yet. Connect a vision AI model and applicable inspection standards to classify this item as Green, Yellow, or Red.");
      }
      setStatus("Inspection photo reviewed.");
    } catch (err) {
      setError(err.message || "Photo analysis failed.");
      setStatus("Photo analysis failed.");
    } finally {
      setPhotoLoading(false);
    }
  }

  function saveFinding() {
    if (!photoAnalysis) {
      setError("Analyse an inspection photo before saving a finding.");
      return;
    }
    setFindings((prev) => [{
      id: Date.now(),
      inspectionItem: currentItem,
      inspectionArea,
      inspectorNotes,
      photoName: inspectionPhoto?.name || "",
      result: photoAnalysis,
    }, ...prev]);
    setInspectionPhoto(null);
    setPhotoPreview("");
    setPhotoAnalysis(null);
    setInspectorNotes("");
    setStatus("Finding saved. Upload or take the next photo for the same item, or select another inspection item.");
  }

  async function finishInspection() {
    setError("");
    try {
      const payload = { brand: "VinQu", generatedAt: new Date().toISOString(), instruction: result, findings };
      const form = new FormData();
      form.append("payload", JSON.stringify(payload));
      const res = await fetch(process.env.NEXT_PUBLIC_BACKEND_URL + "/report/docx", { method: "POST", body: form });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.detail || "Report generation failed.");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "VinQu_Consolidated_Inspection_Report.docx";
      a.click();
      URL.revokeObjectURL(url);
      setStatus("Word report generated.");
    } catch (err) {
      setError(err.message || "Report generation failed.");
    }
  }

  return (
    <main className="page">
      <section className="hero">
        <div className="brand">
          <img src="/logo.jpeg" alt="VinQu logo" className="logo" />
          <div>
            <h1>VinQu</h1>
            <p className="tag">From Evidence to Insight</p>
          </div>
        </div>
        <p className="sub">Upload the inspection instruction, identify inspection items, capture evidence, classify findings, and generate a consolidated Word report.</p>
      </section>

      <section className="grid">
        <div className="card">
          <h2>1) Work instruction / ITP / NOI</h2>
          <label>Upload instruction file</label>
          <input type="file" onChange={onInstructionFile} accept=".pdf,.docx,.txt,.md,.csv,.log,.xlsx" />
          <label>File name / reference</label>
          <input value={fileName} onChange={(e) => setFileName(e.target.value)} />
          <label>Referenced documents for photo review</label>
          <input type="file" multiple onChange={onReferencedDocs} accept=".pdf,.docx,.txt,.md,.csv,.xlsx,.png,.jpg,.jpeg" />
          <p className="hint">Optional: drawings, wiring diagrams, dimensions, electrical test limits, QCP/ITP attachments, or project specifications.</p>
          <button onClick={analyseDocument} disabled={loading}>{loading ? "Analysing..." : "Analyse document"}</button>
          <p className="status">{status}</p>
          {error ? <p className="error">{error}</p> : null}
        </div>

        <div className="card">
          <h2>2) Pre-inspection briefing</h2>
          {!result ? <p className="muted">No instruction analysed yet.</p> : <>
            <div className="resultGrid">
              <Info title="Document / instruction type" value={result.document_type} />
              <Info title="Inspection subject / procedure" value={result.work_instruction_procedure} />
              <Info title="Project / reference" value={result.project_reference || result.project_name || "-"} />
              <Info title="Inspection date window" value={result.inspection_date_window || "-"} />
              <Info title="Location" value={result.inspection_factory_location || result.site_of_inspection || "-"} />
              <Info title="Equipment" value={result.equipment || "-"} />
              <Info title="Scope of work" value={result.scope_of_work || "-"} />
              <Info title="Extraction confidence" value={result.confidence_level || "-"} />
            </div>
            <Block title="Pre-inspection briefing">
              <p>{result.reasoned_summary}</p>
              <p className="hint"><strong>Extraction confidence</strong> means how confidently VinQu extracted fields from the uploaded instruction, not whether the inspection will pass.</p>
            </Block>
            <div className="twoCols">
              <NumberedList title="Inspection activities / tests" items={result.tests_or_checks} />
              <NumberedList title="Inspector duties for this inspection" items={result.inspector_duties} />
              <Block title="Referenced documents"><p>{result.referenced_documents_summary || "No additional referenced documents were uploaded or identified."}</p></Block>
              <List title="Standards" items={result.applicable_standards} />
              <List title="Contacts / parties" items={result.named_contacts} />
              <List title="Missing / uncertain" items={result.missing_information} />
            </div>
            <details className="jsonBox"><summary>Raw JSON / audit view</summary><pre>{JSON.stringify(result, null, 2)}</pre></details>
          </>}
        </div>
      </section>

      <section className="card wide">
        <h2>3) Inspection evidence and findings</h2>
        {!result ? <p className="muted">Analyse an instruction file first.</p> : <div className="workflow">
          <div>
            <label>Current inspection item / test</label>
            <select value={currentItemIndex} onChange={(e) => setCurrentItemIndex(Number(e.target.value))}>
              {inspectionItems.length ? inspectionItems.map((item, i) => <option key={i} value={i}>{i + 1}. {item}</option>) : <option>No test items identified</option>}
            </select>
            <label>Inspection Area / Component</label>
            <input value={inspectionArea} onChange={(e) => setInspectionArea(e.target.value)} placeholder="e.g. terminal block, switch, router, cabinet, relay panel, meter display" />
            <div className="photoButtons">
              <label className="fileButton">Take Inspection Photo<input type="file" accept="image/*" capture="environment" onChange={onPhoto} /></label>
              <label className="fileButton secondary">Upload Inspection Photo<input type="file" accept="image/*" onChange={onPhoto} /></label>
            </div>
            {photoPreview ? <div className="previewBox"><img src={photoPreview} className="preview" alt="Inspection evidence preview" /></div> : null}
            <label>Inspector Notes & Observations</label>
            <textarea value={inspectorNotes} onChange={(e) => setInspectorNotes(e.target.value)} placeholder="VinQu can suggest notes after photo analysis. The inspector can edit or add observations here." />
            <div className="actions">
              <button onClick={analysePhoto} disabled={photoLoading}>{photoLoading ? "Analysing photo..." : "Analyse inspection photo"}</button>
              <button className="ghost" onClick={saveFinding}>Save Current Finding</button>
            </div>
            <div className="actions">
              <button className="ghost" onClick={() => setCurrentItemIndex((i) => Math.min(i + 1, Math.max(inspectionItems.length - 1, 0)))}>Start New Inspection Item</button>
              <button onClick={finishInspection}>Finish Inspection & Generate Word Report</button>
            </div>
          </div>

          <div>
            <h3>Photo analysis result</h3>
            {!photoAnalysis ? <p className="muted">No inspection photo analysed yet.</p> : <div className={"severity " + (photoAnalysis.severity || "manual").toLowerCase()}>
              <div className="severityLabel">{photoAnalysis.status_label}</div>
              <h3>{photoAnalysis.suggested_finding}</h3>
              <p>{photoAnalysis.reasoning}</p>
              <p><strong>Indicative specific reasoning:</strong> {photoAnalysis.standard_reasoning}</p>
              <p><strong>Inspection Area / Component:</strong> {photoAnalysis.inspection_area || inspectionArea || "-"}</p>
              <p><strong>Recommended action:</strong> {photoAnalysis.recommended_action}</p>
            </div>}
            <h3>Saved findings</h3>
            {findings.length ? findings.map((f) => <div className="finding" key={f.id}><strong>{f.result.status_label} — {f.result.suggested_finding}</strong><p>{f.inspectionItem}</p><p>{f.inspectionArea || "No area entered"}</p></div>) : <p className="muted">No saved findings yet.</p>}
          </div>
        </div>}
      </section>
    </main>
  );
}

function Info({ title, value }) { return <div className="info"><div className="t">{title}</div><div className="v">{value || "-"}</div></div>; }
function Block({ title, children }) { return <div className="block"><h3>{title}</h3>{children}</div>; }
function List({ title, items = [] }) { return <Block title={title}>{items && items.length ? <ul>{items.map((item, i) => <li key={i}>{item}</li>)}</ul> : <p className="muted">None</p>}</Block>; }
function NumberedList({ title, items = [] }) { return <Block title={title}>{items && items.length ? <ol>{items.map((item, i) => <li key={i}>{item}</li>)}</ol> : <p className="muted">None</p>}</Block>; }
