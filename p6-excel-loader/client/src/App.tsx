import { useCallback, useMemo, useState } from 'react';
import axios from 'axios';
import { useDropzone } from 'react-dropzone';

interface UploadResponse {
	file: { originalName: string; storedName: string; size: number };
	sheetName: string;
	headers: string[];
	preview: Record<string, any>[];
	totalRows: number;
}

function App() {
	const [uploadInfo, setUploadInfo] = useState<UploadResponse | null>(null);
	const [rows, setRows] = useState<Record<string, any>[]>([]);
	const [mapping, setMapping] = useState({
		wbsCode: '',
		activityId: '',
		activityName: '',
		start: '',
		finish: '',
		duration: '',
		predecessors: ''
	});
	const [project, setProject] = useState({ projectId: 'PRJ1', projectName: 'Imported Project' });
	const [downloading, setDownloading] = useState(false);

	const onDrop = useCallback(async (acceptedFiles: File[]) => {
		if (!acceptedFiles?.length) return;
		const file = acceptedFiles[0];
		const form = new FormData();
		form.append('file', file);
		const { data } = await axios.post<UploadResponse>('http://localhost:4000/api/upload', form, {
			headers: { 'Content-Type': 'multipart/form-data' }
		});
		setUploadInfo(data);
		setRows(data.preview);

		// naive auto-map by header names
		const lower = new Map<string, string>();
		for (const h of data.headers) lower.set(h.toLowerCase(), h);
		const find = (cands: string[]) => {
			for (const c of cands) {
				const v = lower.get(c.toLowerCase());
				if (v) return v;
			}
			return '';
		};
		setMapping({
			wbsCode: find(['WBS Code', 'WBS', 'WBS_ID', 'WBS Code*']),
			activityId: find(['Activity ID', 'ActivityID', 'ID', 'Task ID', 'Activity Number']),
			activityName: find(['Activity Name', 'Name', 'Task Name', 'Description']),
			start: find(['Start', 'Start Date', 'Planned Start', 'Early Start']),
			finish: find(['Finish', 'Finish Date', 'Planned Finish', 'Early Finish']),
			duration: find(['Original Duration', 'Duration', 'Planned Duration']),
			predecessors: find(['Predecessors', 'Predecessor ID', 'Logic', 'Links'])
		});
	}, []);

	const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop, multiple: false, accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'], 'text/csv': ['.csv'] } });

	const requiredOk = useMemo(() => mapping.activityId && mapping.activityName, [mapping]);

	const downloadP6Workbook = async () => {
		setDownloading(true);
		try {
			const { data } = await axios.post('http://localhost:4000/api/generate/p6-excel', {
				project,
				mapping,
				rows
			}, { responseType: 'blob' });
			const url = window.URL.createObjectURL(new Blob([data]));
			const a = document.createElement('a');
			a.href = url;
			a.download = `p6_loader_${Date.now()}.xlsx`;
			a.click();
			window.URL.revokeObjectURL(url);
		} finally {
			setDownloading(false);
		}
	};

	return (
		<div style={{ maxWidth: 1100, margin: '0 auto', padding: 24, fontFamily: 'Inter, system-ui, Arial' }}>
			<h2>Primavera P6 Excel Loader</h2>
			<p>Upload an Excel file, map columns, preview, and export a P6-friendly workbook for import (Activities & Relationships).</p>

			<div {...getRootProps()} style={{ border: '2px dashed #888', padding: 24, textAlign: 'center', borderRadius: 8, background: isDragActive ? '#fafafa' : 'transparent' }}>
				<input {...getInputProps()} />
				{isDragActive ? <p>Drop the file here...</p> : <p>Drag & drop an Excel file here, or click to select</p>}
			</div>

			{uploadInfo && (
				<div style={{ marginTop: 24 }}>
					<h3>File Parsed</h3>
					<div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
						<div>
							<label>Project ID</label><br />
							<input value={project.projectId} onChange={e => setProject(p => ({ ...p, projectId: e.target.value }))} />
						</div>
						<div>
							<label>Project Name</label><br />
							<input style={{ width: '100%' }} value={project.projectName} onChange={e => setProject(p => ({ ...p, projectName: e.target.value }))} />
						</div>
						<div>
							<label>WBS Code</label><br />
							<select value={mapping.wbsCode} onChange={e => setMapping(m => ({ ...m, wbsCode: e.target.value }))}>
								<option value="">-- optional --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Activity ID</label><br />
							<select value={mapping.activityId} onChange={e => setMapping(m => ({ ...m, activityId: e.target.value }))}>
								<option value="">-- select --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Activity Name</label><br />
							<select value={mapping.activityName} onChange={e => setMapping(m => ({ ...m, activityName: e.target.value }))}>
								<option value="">-- select --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Start (optional)</label><br />
							<select value={mapping.start} onChange={e => setMapping(m => ({ ...m, start: e.target.value }))}>
								<option value="">-- optional --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Finish (optional)</label><br />
							<select value={mapping.finish} onChange={e => setMapping(m => ({ ...m, finish: e.target.value }))}>
								<option value="">-- optional --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Duration (optional)</label><br />
							<select value={mapping.duration} onChange={e => setMapping(m => ({ ...m, duration: e.target.value }))}>
								<option value="">-- optional --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
						<div>
							<label>Predecessors (optional)</label><br />
							<select value={mapping.predecessors} onChange={e => setMapping(m => ({ ...m, predecessors: e.target.value }))}>
								<option value="">-- optional --</option>
								{uploadInfo.headers.map(h => <option key={h} value={h}>{h}</option>)}
							</select>
						</div>
					</div>

					<div style={{ marginTop: 24 }}>
						<button disabled={!requiredOk || downloading} onClick={downloadP6Workbook}>{downloading ? 'Generating...' : 'Download P6 Workbook (.xlsx)'}</button>
					</div>

					<div style={{ marginTop: 24 }}>
						<h3>Preview ({uploadInfo.totalRows} rows, showing up to 100)</h3>
						<div style={{ overflow: 'auto', maxHeight: 360, border: '1px solid #eee' }}>
							<table style={{ borderCollapse: 'collapse', width: '100%' }}>
								<thead>
									<tr>
										{uploadInfo.headers.map(h => (
											<th key={h} style={{ position: 'sticky', top: 0, background: '#fafafa', borderBottom: '1px solid #ddd', textAlign: 'left', padding: '6px 8px' }}>{h}</th>
										))}
									</tr>
								</thead>
								<tbody>
									{rows.map((r, idx) => (
										<tr key={idx}>
											{uploadInfo.headers.map(h => (
												<td key={h} style={{ borderBottom: '1px solid #f0f0f0', padding: '6px 8px', whiteSpace: 'nowrap' }}>{String(r[h] ?? '')}</td>
											))}
										</tr>
									))}
								</tbody>
							</table>
						</div>
					</div>
				</div>
			)}
		</div>
	);
}

export default App;