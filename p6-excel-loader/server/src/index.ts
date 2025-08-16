import express from 'express';
import cors from 'cors';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import XLSX from 'xlsx';
import { z } from 'zod';

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const upload = multer({
	storage: multer.diskStorage({
		destination: (req, file, cb) => {
			const dir = path.join(process.cwd(), 'uploads');
			fs.mkdirSync(dir, { recursive: true });
			cb(null, dir);
		},
		filename: (req, file, cb) => {
			const ext = path.extname(file.originalname);
			cb(null, `${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`);
		}
	}),
	limits: { fileSize: 20 * 1024 * 1024 }
});

app.get('/api/health', (_req, res) => {
	res.json({ ok: true });
});

// Upload Excel and parse first sheet to JSON (array of rows)
app.post('/api/upload', upload.single('file'), (req, res) => {
	try {
		if (!req.file) {
			return res.status(400).json({ error: 'No file uploaded' });
		}
		const workbook = XLSX.read(fs.readFileSync(req.file.path));
		const sheetName = workbook.SheetNames[0];
		const sheet = workbook.Sheets[sheetName];
		const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });
		const headers = Object.keys(rows[0] || {});
		res.json({
			file: { originalName: req.file.originalname, storedName: req.file.filename, size: req.file.size },
			sheetName,
			headers,
			preview: rows.slice(0, 100),
			totalRows: rows.length
		});
	} catch (e: any) {
		res.status(500).json({ error: e?.message || 'Failed to parse Excel' });
	}
});

// Generate P6-friendly Activities/Relationships workbook based on a mapping
const GenerateRequestSchema = z.object({
	project: z.object({
		projectId: z.string().min(1),
		projectName: z.string().min(1)
	}),
	mapping: z.object({
		wbsCode: z.string().optional().default(''),
		activityId: z.string().min(1),
		activityName: z.string().min(1),
		start: z.string().optional().default(''),
		finish: z.string().optional().default(''),
		duration: z.string().optional().default(''),
		predecessors: z.string().optional().default('')
	}),
	rows: z.array(z.record(z.any()))
});

app.post('/api/generate/p6-excel', (req, res) => {
	try {
		const parsed = GenerateRequestSchema.parse(req.body);
		const { mapping, rows, project } = parsed;

		const activities: any[] = [];
		const relationships: any[] = [];

		for (const row of rows) {
			const wbs = String(row[mapping.wbsCode || ''] ?? '').trim();
			const actId = String(row[mapping.activityId] ?? '').trim();
			const actName = String(row[mapping.activityName] ?? '').trim();
			if (!actId || !actName) continue;

			const startVal = mapping.start ? String(row[mapping.start] ?? '').trim() : '';
			const finishVal = mapping.finish ? String(row[mapping.finish] ?? '').trim() : '';
			const durationVal = mapping.duration ? String(row[mapping.duration] ?? '').trim() : '';

			activities.push({
				'Project ID': project.projectId,
				'WBS Code': wbs,
				'Activity ID': actId,
				'Activity Name': actName,
				'Activity Type': 'Task Dependent',
				'Original Duration': durationVal,
				'Remaining Duration': durationVal,
				'Start': startVal,
				'Finish': finishVal,
				'Calendar': ''
			});

			const predsText = mapping.predecessors ? String(row[mapping.predecessors] ?? '').trim() : '';
			if (predsText) {
				for (const token of predsText.split(/[,;]+/).map((t: string) => t.trim()).filter(Boolean)) {
					// Parse forms like: A100, A100FS, A100FS+2d, A100SS-1d
					const m = token.match(/^(?<pred>[A-Za-z0-9_-]+)(?:(?<type>FS|FF|SS|SF))?(?:(?<lagSign>[+-])(?<lagVal>[0-9]+)(?<lagUnit>[dhw])?)?$/i);
					if (!m || !m.groups) continue;
					const pred = m.groups['pred'];
					const relType = (m.groups['type'] || 'FS').toUpperCase();
					const sign = m.groups['lagSign'] || '+';
					const lagVal = m.groups['lagVal'] || '0';
					const lagUnit = (m.groups['lagUnit'] || 'd').toLowerCase();
					const lagDays = lagUnit === 'h' ? Number(lagVal) / 8 : (lagUnit === 'w' ? Number(lagVal) * 5 : Number(lagVal));
					const lag = (sign === '-' ? -lagDays : lagDays).toString();
					relationships.push({
						'Project ID': project.projectId,
						'Predecessor ID': pred,
						'Successor ID': actId,
						'Type': relType,
						'Lag (days)': lag
					});
				}
			}
		}

		const wb = XLSX.utils.book_new();
		const actSheet = XLSX.utils.json_to_sheet(activities);
		XLSX.utils.book_append_sheet(wb, actSheet, 'Activities');
		if (relationships.length > 0) {
			const relSheet = XLSX.utils.json_to_sheet(relationships);
			XLSX.utils.book_append_sheet(wb, relSheet, 'Relationships');
		}

		const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
		res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		res.setHeader('Content-Disposition', `attachment; filename="p6_loader_${Date.now()}.xlsx"`);
		res.send(buf);
	} catch (e: any) {
		res.status(400).json({ error: e?.message || 'Invalid input' });
	}
});

const PORT = Number(process.env.PORT) || 4000;
app.listen(PORT, () => {
	console.log(`Server listening on http://localhost:${PORT}`);
});