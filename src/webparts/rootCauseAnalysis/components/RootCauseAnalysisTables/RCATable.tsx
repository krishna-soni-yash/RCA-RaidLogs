import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { DetailsList, DetailsRow, IDetailsRowProps, IColumn, SelectionMode, CheckboxVisibility, PrimaryButton, DefaultButton, Dialog, DialogType, IconButton, mergeStyleSets } from '@fluentui/react';
import RCAForm from '../RootCauseAnalysisForms/RCAForm';
import historyStyles from './RCATableHistory.module.scss';
import { RCACOLUMNS, SubSiteListNames } from '../../../../common/Constants';
import { IRCAList } from '../../../../models/IRCAList';
import { GenericService } from '../../../../services/GenericServices';
import IGenericService from '../../../../services/IGenericServices';
import { getRCAItems, RCARepository } from '../../../../repositories/RCARepository';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { exportRowsToExcel } from '../../../../common/excelExport';
import { formatResponsibilityValue, formatDateMMDDYYYY } from '../../../../common/exportHelpers';
import IRCARepository from '../../../../repositories/repositoriesInterface/IRCARepository';
export interface IColumnConfig {
	key: string;
	name: string;
	fieldName: string;
	minWidth: number;
	maxWidth: number;
	isResizable?: boolean;
	onRender?: (item: any) => React.ReactNode;
}



interface RCATableProps {
	// optional preset columns; if omitted columns will be generated from the first item
	columns?: IColumn[];
	// optional compact mode
	compact?: boolean;
	// optional className for styling
	className?: string;
	// web part context
	context: WebPartContext;
}


const PAGE_SIZE = 8;

const classNames = mergeStyleSets({
	container: {
		padding: 12,
		background: '#ffffff',
		borderRadius: 6,
		boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
	},
	detailsList: {
		// tweaks to header and row padding for compact, aligned look
		selectors: {
			// ensure the list uses full width and remove content wrapper extra right padding
			// '.ms-DetailsList': {
			// 	width: '100%'
			// },
			// '.ms-DetailsList-contentWrapper': {
			// 	paddingRight: 0
			// },
			// '.ms-DetailsHeader': {
			// 	padding: '8px 8px', // less right padding
			// 	background: '#f3f2f1',
			// 	borderRadius: '6px 6px 0 0'
			// },
			// '.ms-DetailsHeader-cell': {
			// 	paddingRight: 8
			// },
			// '.ms-DetailsRow': {
			// 	padding: '6px 8px', // reduce right padding
			// 	alignItems: 'center'
			// },
			// '.ms-DetailsRow-cell': {
			// 	display: 'flex',
			// 	alignItems: 'center',
			// 	paddingRight: 8
			// }
		}
	},
	rowWrapper: {
		display: 'flex',
		alignItems: 'center'
	},
	expandButton: {
		width: 36,
		display: 'flex',
		justifyContent: 'center',
		alignItems: 'center'
	},
	subtableContainer: {
		/* reduce left/right padding and make subtable visually tighter to match main list */
		padding: '8px 12px 12px 56px',
		background: '#ffffff',
		borderLeft: '4px solid #0078d4',
		borderBottom: '1px solid #eee',
		borderRight: '1px solid #eee',
		borderRadius: '0 0 6px 6px',
		boxShadow: 'inset 0 1px 0 rgba(0,0,0,0.02)'
	},
	subDetailsList: {
		selectors: {
		}
	},
	paginationBar: {
		display: 'flex',
		justifyContent: 'space-between',
		alignItems: 'center',
		marginTop: 12,
		padding: '8px 12px',
		background: '#f3f2f1',
		borderRadius: 4
	},
	paginationControls: {
		display: 'flex',
		gap: 8
	},
	historyContainer: {
		maxHeight: 560,
		overflowY: 'auto',
		padding: '4px 4px 4px 0',
		background: '#fafafa'
	},
	historyModalContent: {
		background: '#ffffff',
		borderRadius: 8,
		boxShadow: '0 4px 20px rgba(0,0,0,0.2)',
		width: '100%',
		maxHeight: '90vh',
		display: 'flex',
		flexDirection: 'column'
	},
	historyModalHeader: {
		padding: '18px 22px',
		borderBottom: '1px solid #e0e0e0',
		display: 'flex',
		justifyContent: 'space-between',
		alignItems: 'center'
	},
	historyModalTitle: {
		fontSize: 20,
		color: '#333',
		margin: 0,
		fontWeight: 600
	},
	historyModalBody: {
		padding: '16px 20px',
		overflowY: 'auto',
		flex: 1,
		background: '#fafafa'
	},
	historySummary: {
		background: '#dff6dd',
		border: '1px solid #a7e5a3',
		borderRadius: 6,
		padding: '10px 12px',
		marginBottom: 12,
		fontSize: 13,
		fontWeight: 600,
		color: '#1f7a1f'
	},
	historyCard: {
		border: '1px solid #e0e0e0',
		borderLeft: '4px solid #d0d0d0',
		borderRadius: 8,
		padding: 12,
		marginBottom: 12,
		background: '#ffffff',
		boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
		selectors: {
			':hover': {
				boxShadow: '0 2px 6px rgba(0,0,0,0.15)'
			}
		}
	},
	historyCardLatest: {
		borderLeftColor: '#0078d4',
		background: '#f0f8ff'
	},
	historyCardHeader: {
		display: 'flex',
		justifyContent: 'space-between',
		alignItems: 'center',
		marginBottom: 8,
		paddingBottom: 8,
		borderBottom: '1px solid #eceff1'
	},
	historyVersionBadge: {
		fontSize: 12,
		fontWeight: 700,
		padding: '4px 8px',
		borderRadius: 999,
		background: '#e8f4fd',
		color: '#005a9e'
	},
	historyMetaText: {
		fontSize: 12,
		color: '#5c6773'
	},
	historySection: {
		marginTop: 10
	},
	historySectionTitle: {
		fontSize: 12,
		fontWeight: 700,
		color: '#323130',
		textTransform: 'uppercase',
		letterSpacing: 0.3,
		marginBottom: 6,
		paddingBottom: 4,
		borderBottom: '1px solid #f0f0f0'
	},
	historyFieldGrid: {
		display: 'grid',
		gridTemplateColumns: '240px 1fr',
		columnGap: 10,
		rowGap: 6,
		fontSize: 12
	},
	historyFieldLabel: {
		fontWeight: 600,
		color: '#323130'
	},
	historyFieldValue: {
		color: '#201f1e',
		wordBreak: 'break-word'
	}
});

const CAUSAL_ACTION_SEGMENTS = [
	{ suffix: 'Correction', label: 'Correction' },
	{ suffix: 'Corrective', label: 'Corrective Action' },
	{ suffix: 'Preventive', label: 'Preventive Action' }
];

	const CAUSAL_EXPORT_HEADERS = [
		'Problem statement (Causal Analysis Trigger)',
		'Cause Category',
		'Source',
		'Priority',
		'Related Metric (if any)',
		'Related Sub Metric (if any)',
		'Cause(s)',
		'Root Cause(s)',
		'Root Cause Analysis Technique Used and Reference (if any)',
		'Type of Action',
		'Action Plan (Correction)',
		'Responsibility (Correction)',
		'Planned Closure Date (Correction)',
		'Actual Closure Date (Correction)',
		'Action Plan (Corrective Action)',
		'Responsibility (Corrective Action)',
		'Planned Closure Date (Corrective Action)',
		'Actual Closure Date (Corrective Action)',
		'Action Plan (Preventive Action)',
		'Responsibility (Preventive Action)',
		'Planned Closure Date (Preventive Action)',
		'Actual Closure Date (Preventive Action)',
		'Performance before action plan',
		'Performance after action plan',
		'Quantitative / Statistical effectiveness',
		'Remarks'
	];

	const buildCausalExportRow = (item: Partial<IRCAList>): Record<string, any> => {
		// Format Type of Action - handle array or comma-separated string
		let typeOfActionStr = '';
		if (Array.isArray(item.RCATypeOfAction)) {
			typeOfActionStr = item.RCATypeOfAction.map((t: any) => String(t).trim()).filter(Boolean).join('; ');
		} else if (typeof item.RCATypeOfAction === 'string' && item.RCATypeOfAction.trim().length) {
			typeOfActionStr = item.RCATypeOfAction.split(',').map((s: string) => s.trim()).filter(Boolean).join('; ');
		}

		const row: Record<string, any> = {
			'Problem statement (Causal Analysis Trigger)': item.LinkTitle ?? '',
			'Cause Category': item.CauseCategory ?? '',
			'Source': item.RCASource ?? '',
			'Priority': item.RCAPriority ?? '',
			'Related Metric (if any)': item.RelatedMetric ?? '',
			'Related Sub Metric (if any)': item.RelatedSubMetric ?? '',
			'Cause(s)': item.Cause ?? '',
			'Root Cause(s)': item.RootCause ?? '',
			'Root Cause Analysis Technique Used and Reference (if any)': item.RCATechniqueUsedAndReference ?? '',
			'Type of Action': typeOfActionStr
		};

		CAUSAL_ACTION_SEGMENTS.forEach((segment) => {
			const planKey = `ActionPlan${segment.suffix}` as keyof Partial<IRCAList>;
			const responsibilityKey = `Responsibility${segment.suffix}` as keyof Partial<IRCAList>;
			const plannedKey = `PlannedClosureDate${segment.suffix}` as keyof Partial<IRCAList>;
			const actualKey = `ActualClosureDate${segment.suffix}` as keyof Partial<IRCAList>;

			row[`Action Plan (${segment.label})`] = item[planKey] ?? '';
			row[`Responsibility (${segment.label})`] = formatResponsibilityValue(item[responsibilityKey]);
			row[`Planned Closure Date (${segment.label})`] = formatDateMMDDYYYY(item[plannedKey]);
			row[`Actual Closure Date (${segment.label})`] = formatDateMMDDYYYY(item[actualKey]);
		});

		row['Performance before action plan'] = item.PerformanceBeforeActionPlan ?? '';
		row['Performance after action plan'] = item.PerformanceAfterActionPlan ?? '';
		row['Quantitative / Statistical effectiveness'] = item.QuantitativeOrStatisticalEffecti ?? '';
		row['Remarks'] = item.Remarks ?? '';
		return row;
	};

	const getModifiedTimestamp = (entry: Partial<IRCAList> | undefined): number => {
		if (!entry) return 0;
		const raw: any = (entry as any).Modified ?? (entry as any).modified ?? (entry as any).LastModified ?? (entry as any).lastModified;
		if (raw instanceof Date) {
			return raw.getTime();
		}
		if (typeof raw === 'string' || typeof raw === 'number') {
			const parsed = new Date(raw);
			const time = parsed.getTime();
			return isNaN(time) ? 0 : time;
		}
		return 0;
	};

const RCATable: React.FC<RCATableProps> = ({ columns, compact, context, className }) => {
	// prefer passed columns, then RCACOLUMNS, then fallback dynamic columns
	const cols = columns && columns.length ? columns : RCACOLUMNS;

	// local state so we can append new items added via the form dialog
	const [localItems, setLocalItems] = useState<any[]>([]);
	const [isDialogOpen, setIsDialogOpen] = useState(false);
	const [RCAItems, setRCAItems] = useState<Partial<IRCAList>[]>([]);

	// editing state
	const [selectedItem, setSelectedItem] = useState<Partial<IRCAList> | null>(null);
	const [isEditing, setIsEditing] = useState<boolean>(false);
	const [isHistoryDialogOpen, setIsHistoryDialogOpen] = useState<boolean>(false);
	const [isHistoryLoading, setIsHistoryLoading] = useState<boolean>(false);
	const [historyItemTitle, setHistoryItemTitle] = useState<string>('');
	const [historyVersions, setHistoryVersions] = useState<any[]>([]);

	const openDialog = () => setIsDialogOpen(true);
	const closeDialog = () => {
		setIsDialogOpen(false);
		setSelectedItem(null);
		setIsEditing(false);
	};

	const fetchRCAItems = useCallback(async () => {
		const genericServiceInstance: IGenericService = new GenericService(undefined, context);
		genericServiceInstance.init(undefined, context);
		const rcaRepo: IRCARepository = new RCARepository(genericServiceInstance);
		rcaRepo.setService(genericServiceInstance);
		const items = await getRCAItems(true, context);
		const sortedItems = Array.isArray(items)
			? [...items].sort((a: Partial<IRCAList>, b: Partial<IRCAList>) => {
				const diff = getModifiedTimestamp(b) - getModifiedTimestamp(a);
				if (diff !== 0) return diff;
				const bId = Number((b as any)?.ID ?? (b as any)?.Id ?? 0);
				const aId = Number((a as any)?.ID ?? (a as any)?.Id ?? 0);
				return bId - aId;
			})
			: items;
		// newest items first so users see the latest updates immediately
		setRCAItems(sortedItems);
	}, [context]);

	// helper: map repository item (IRCAList) to RCAForm initialData shape
	const mapRepoItemToForm = (it: any): any => {
		if (!it) return {};
		const form: any = {};
		const parsePeopleValues = (value: any): string[] => {
			if (!value) return [];
			if (Array.isArray(value)) {
				return value
					.map((entry: any) => String(entry || '').trim())
					.filter((entry: string) => entry.length > 0);
			}
			if (typeof value === 'string') {
				return value
					.split(/; ?/)
					.map((entry: string) => entry.trim())
					.filter((entry: string) => entry.length > 0);
			}
			return [];
		};
		form.problemStatement = it.ProblemStatement || it.LinkTitle || '';
		form.causeCategory = it.CauseCategory || '';
		form.source = it.RCASource || '';
		form.priority = it.RCAPriority || '';
		form.relatedMetric = it.RelatedMetric || '';
		form.causes = it.Cause || '';
		form.rootCauses = it.RootCause || '';
		form.analysisTechnique = it.RCATechniqueUsedAndReference || '';
		// action types -> array
		form.actionType = it.RCATypeOfAction ? (typeof it.RCATypeOfAction === 'string' ? it.RCATypeOfAction.split(',').map((s: string) => s.trim()).filter(Boolean) : it.RCATypeOfAction) : [];

		// build actionDetails for each known action type
		const details: Record<string, any> = {};
		const actionKeys = form.actionType.length ? form.actionType : ['Correction', 'Corrective Action', 'Preventive Action'];

		actionKeys.forEach((act: string) => {
			let suffix = '';
			const lower = (act || '').toString().toLowerCase();
			if (lower.indexOf('correction') !== -1) suffix = 'Correction';
			else if (lower.indexOf('corrective') !== -1) suffix = 'Corrective';
			else if (lower.indexOf('preventive') !== -1) suffix = 'Preventive';
			else suffix = act.replace(/\s+/g, '');

			details[act] = {
				actionPlan: it[`ActionPlan${suffix}`] || '',
				// Normalize multi-people picker values into id|value strings
				responsibility: parsePeopleValues(it[`Responsibility${suffix}`]),
				plannedClosureDate: it[`PlannedClosureDate${suffix}`] ? new Date(it[`PlannedClosureDate${suffix}`]) : undefined,
				actualClosureDate: it[`ActualClosureDate${suffix}`] ? new Date(it[`ActualClosureDate${suffix}`]) : undefined
			};
		});

		form.actionDetails = details;
		form.performanceBefore = it.PerformanceBeforeActionPlan || '';
		form.performanceAfter = it.PerformanceAfterActionPlan || '';
		form.quantitativeEffectiveness = it.QuantitativeOrStatisticalEffecti || '';
		form.remarks = it.Remarks || '';
		form.relatedSubMetric = it.RelatedSubMetric || '';
		form.attachments = (it.attachments && Array.isArray(it.attachments)) ? it.attachments.map((a: any) => ({
			FileName: a.FileName || a.fileName || '',
			ServerRelativeUrl: a.ServerRelativeUrl || a.Url || a.FileRef || ''
		})) : [];
		// preserve id for editing context
		form.__repoId = it.ID ?? it.Id ?? it.Id;
		return form;
	};

	const handleFormSubmit = async (data: any) => {
		// if editing, refresh remote list (saved by RCAForm) to reflect changes
		if (isEditing) {
			// fetchRCAItems will refresh displayed rows
			await fetchRCAItems();
			closeDialog();
			return;
		}else {
			// new item added; append to localItems
			await fetchRCAItems();
		}

		// adding new local item (existing behaviour)
		const normalized = { ...data };
		if (normalized.plannedClosureDate instanceof Date) {
			normalized.plannedClosureDate = normalized.plannedClosureDate.toLocaleDateString();
		}
		if (normalized.actualClosureDate instanceof Date) {
			normalized.actualClosureDate = normalized.actualClosureDate.toLocaleDateString();
		}
		setLocalItems((prev) => [...prev, normalized]);
		closeDialog();
	};

	const openVersionHistory = async (item: Partial<IRCAList>): Promise<void> => {
		const rawId: any = (item as any)?.ID ?? (item as any)?.Id;
		const itemId = Number(rawId);
		if (!itemId || isNaN(itemId)) return;

		setHistoryItemTitle(String((item as any)?.LinkTitle ?? `RCA ${itemId}`));
		setIsHistoryLoading(true);
		setIsHistoryDialogOpen(true);

		try {
			const genericServiceInstance: IGenericService = new GenericService(undefined, context);
			genericServiceInstance.init(undefined, context);

			const versions = await genericServiceInstance.getVersionHistory<any>({
				context,
				listTitle: SubSiteListNames.RootCauseAnalysis,
				itemId,
				select: [
					'VersionLabel', 'Created', 'Modified', 'CheckInComment', 'Editor/Title', 'Editor/EMail',
					'LinkTitle',
					'CauseCategory',
					'RCASource',
					'RCAPriority',
					'RelatedMetric',
					'RelatedSubMetric',
					'Cause',
					'RootCause',
					'RCATechniqueUsedAndReference',
					'RCATypeOfAction',
					'ActionPlanCorrection',
					'ResponsibilityCorrection',
					'PlannedClosureDateCorrection',
					'ActualClosureDateCorrection',
					'ActionPlanCorrective',
					'ResponsibilityCorrective',
					'PlannedClosureDateCorrective',
					'ActualClosureDateCorrective',
					'ActionPlanPreventive',
					'ResponsibilityPreventive',
					'PlannedClosureDatePreventive',
					'ActualClosureDatePreventive',
					'PerformanceBeforeActionPlan',
					'PerformanceAfterActionPlan',
					'QuantitativeOrStatisticalEffecti',
					'Remarks'
				],
				expand: ['Editor']
			});

			setHistoryVersions(Array.isArray(versions) ? versions : []);
		} catch (error) {
			console.error('Error fetching RCA version history:', error);
			setHistoryVersions([]);
		} finally {
			setIsHistoryLoading(false);
		}
	};

	const textValue = (value: any): string => {
		if (value === null || value === undefined) return '-';
		if (Array.isArray(value)) {
			const normalized = value.map((entry: any) => String(entry ?? '').trim()).filter(Boolean);
			return normalized.length ? normalized.join('; ') : '-';
		}
		if (typeof value === 'string') {
			const trimmed = value.trim();
			return trimmed.length ? trimmed : '-';
		}
		if (typeof value === 'number' || typeof value === 'boolean') {
			return String(value);
		}
		if (typeof value === 'object') {
			const title = (value as any).Title ?? (value as any).title;
			if (title) return String(title);
		}
		return String(value);
	};

	const actionTypeText = (version: any): string => {
		const raw = version?.RCATypeOfAction;
		if (Array.isArray(raw)) {
			return textValue(raw);
		}
		if (typeof raw === 'string' && raw.trim().length) {
			return raw.split(',').map((x: string) => x.trim()).filter(Boolean).join('; ');
		}
		return '-';
	};

	useEffect(() => {
		void fetchRCAItems();
	}, [fetchRCAItems]);

	// If an RCAId (or variants) query parameter is present, open that item in the edit dialog
	useEffect(() => {
		const tryOpenFromQuery = async (): Promise<void> => {
			try {
				const params = new URLSearchParams(window.location.search);
				const raw = params.get('RCAId') || params.get('RcaId') || params.get('rcaid') || params.get('RCAid') || params.get('rcaId');
				if (!raw) return;

				const value = raw;
				let found: any | undefined = undefined;

				for (let idx = 0; idx < RCAItems.length; idx++) {
					const it = RCAItems[idx];
					if (String(it.ID) === value) {
						found = it;
						break;
					}
				}

				// If not found in current list, try fetching fresh (in case items not yet loaded)
				if (!found && !isNaN(Number(value))) {
					try {
						const refreshed = await getRCAItems(false, context);
						if (refreshed && refreshed.length > 0) {
							for (let idx = 0; idx < refreshed.length; idx++) {
								const it = refreshed[idx];
								if (String(it.ID) === value) { found = it; break; }
							}
							// update local state so UI shows the item in list
							setRCAItems(refreshed);
						}
					} catch (e) {
						// ignore
					}
				}

				if (found) {
					setSelectedItem(found);
					setIsEditing(true);
					setIsDialogOpen(true);

					// Remove query param so modal doesn't reopen on refresh
					const newUrl = new URL(window.location.href);
					newUrl.searchParams.delete('RCAId');
					newUrl.searchParams.delete('RcaId');
					newUrl.searchParams.delete('rcaid');
					newUrl.searchParams.delete('RCAid');
					newUrl.searchParams.delete('rcaId');
					window.history.replaceState(null, '', newUrl.toString());
				}
			} catch (err) {
				console.error('Error opening RCA item from query param:', err);
			}
		};

		if (RCAItems && RCAItems.length > 0) {
			void tryOpenFromQuery();
		}
	}, [RCAItems, context]);

	// edit column prepended to columns
	const allowedMainFieldNames = new Set(['LinkTitle', 'CauseCategory', 'RCAPriority', 'RootCause', 'RCATypeOfAction']);
	const mainListColumns: IColumn[] = (cols as IColumn[]).filter((column: IColumn) => allowedMainFieldNames.has(String(column.fieldName || '')));

	const displayedColumns: IColumn[] = [
		{
			key: 'edit',
			name: '',
			fieldName: 'edit',
			minWidth: 72,
			maxWidth: 72,
			isResizable: false,
			onRender: (item: any) => (
				<div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
					<IconButton
						menuIconProps={{ iconName: '' }}
						iconProps={{ iconName: 'Edit' }}
						title="Edit"
						ariaLabel="Edit"
						onClick={() => {
							// open dialog with mapped initial data
							setSelectedItem(item);
							setIsEditing(true);
							setIsDialogOpen(true);
						}}
					/>
					<IconButton
						menuIconProps={{ iconName: '' }}
						iconProps={{ iconName: 'History' }}
						title="Version History"
						ariaLabel="Version History"
						styles={{
							root: { color: '#107C10' },
							rootHovered: { color: '#107C10', background: '#e6f4ea' },
						}}
						onClick={() => { void openVersionHistory(item); }}
					/>
				</div>
			)
		},
		...mainListColumns
	];

	// expanded rows state (store string keys derived from each item)
	const [expandedKeys, setExpandedKeys] = useState<string[]>([]);
	const [currentPage, setCurrentPage] = useState<number>(1);
	const keyForItem = (item: any) =>
		String(item?.ID ?? item?.id ?? item?.key ?? item?.__repoId ?? item?.LinkTitle ?? JSON.stringify(item).slice(0, 40));

	const toggleExpand = (item: any) => {
		const k = keyForItem(item);
		setExpandedKeys((prev) => {
			const exists = prev.indexOf(k) !== -1;
			if (exists) return prev.filter((x) => x !== k);
			return [...prev, k];
		});
	};

	// prune expanded keys when source items change so new data stays collapsed by default
	useEffect(() => {
		const sourceItems = (RCAItems && RCAItems.length > 0) ? RCAItems : (localItems && localItems.length > 0 ? localItems : []);
		if (!sourceItems || sourceItems.length === 0) {
			if (expandedKeys.length > 0) {
				setExpandedKeys([]);
			}
			return;
		}
		const validKeys = new Set(sourceItems.map((it: any) => keyForItem(it)));
		const filtered = expandedKeys.filter((key) => validKeys.has(key));
		if (filtered.length !== expandedKeys.length) {
			setExpandedKeys(filtered);
		}
	}, [RCAItems, localItems, expandedKeys]);

	useEffect(() => {
		const totalPages = Math.max(1, Math.ceil((RCAItems?.length ?? 0) / PAGE_SIZE));
		setCurrentPage((prev) => {
			if (prev > totalPages) return totalPages;
			if (prev < 1) return 1;
			return prev;
		});
	}, [RCAItems]);

	const totalPages = Math.max(1, Math.ceil((RCAItems?.length ?? 0) / PAGE_SIZE));
	const safeCurrentPage = Math.min(Math.max(currentPage, 1), totalPages);
	const paginatedItems = React.useMemo(() => {
		if (!RCAItems || RCAItems.length === 0) return [];
		const start = (safeCurrentPage - 1) * PAGE_SIZE;
		return RCAItems.slice(start, start + PAGE_SIZE);
	}, [RCAItems, safeCurrentPage]);

	const canGoPrevious = safeCurrentPage > 1;
	const canGoNext = safeCurrentPage < totalPages;

	const handleExportCausalAnalysis = (): void => {
		const rows = (RCAItems ?? []).map(buildCausalExportRow);
		exportRowsToExcel({
			rows,
			headers: CAUSAL_EXPORT_HEADERS,
			sheetName: 'Causal Analysis',
			fileName: 'CausalAnalysis'
		});
	};

	// render a compact 3-row table for the action types under a parent row
	const renderActionSubTable = (it: any) => {
		if (!it) return null;

		// derive action types from the item (string or array). fallback to the default set.
		let actionTypes: string[] = [];
		if (Array.isArray(it?.RCATypeOfAction)) actionTypes = it.RCATypeOfAction.map((t: any) => String(t).trim()).filter(Boolean);
		else if (typeof it?.RCATypeOfAction === 'string' && it.RCATypeOfAction.trim().length) {
			actionTypes = it.RCATypeOfAction.split(',').map((s: string) => s.trim()).filter(Boolean);
		} else {
			actionTypes = ['Correction', 'Corrective Action', 'Preventive Action'];
		}

		// build rows dynamically based on detected action types and suffix mapping
		const rows = actionTypes.map((act: string, idx: number) => {
			const lower = (act || '').toString().toLowerCase();
			let suffix = '';
			if (lower.indexOf('correction') !== -1) suffix = 'Correction';
			else if (lower.indexOf('corrective') !== -1) suffix = 'Corrective';
			else if (lower.indexOf('preventive') !== -1) suffix = 'Preventive';
			else suffix = act.replace(/\s+/g, '');

			const actionPlan = it[`ActionPlan${suffix}`] ?? '';
			const responsibility = formatResponsibilityValue(it[`Responsibility${suffix}`]);
			const planned = formatDateMMDDYYYY(it[`PlannedClosureDate${suffix}`] ?? '');
			const actual = formatDateMMDDYYYY(it[`ActualClosureDate${suffix}`] ?? '');

			return {
				key: `${suffix}-${idx}`,
				type: act,
				actionPlan,
				responsibility,
				planned,
				actual
			};
		});

		const subColumns: IColumn[] = [
			{ key: 'type', name: 'Type of Action', fieldName: 'type', minWidth: 100, maxWidth: 140, isResizable: true },
			{ key: 'actionPlan', name: 'Action Plan', fieldName: 'actionPlan', minWidth: 180, maxWidth: 360, isResizable: true },
			{ key: 'responsibility', name: 'Responsibility', fieldName: 'responsibility', minWidth: 120, maxWidth: 180, isResizable: true },
			{ key: 'planned', name: 'Planned', fieldName: 'planned', minWidth: 90, maxWidth: 130, isResizable: true },
			{ key: 'actual', name: 'Actual', fieldName: 'actual', minWidth: 90, maxWidth: 130, isResizable: true }
		];

		return (
			<div className={classNames.subtableContainer}>
				<DetailsList
					items={rows}
					columns={subColumns}
					selectionMode={SelectionMode.none}
					checkboxVisibility={CheckboxVisibility.hidden}
					compact={true}
					className={classNames.subDetailsList}
					setKey={`subtable-${it.ID ?? it.__repoId ?? Math.random()}`}
					isHeaderVisible={true}
				/>
			</div>
		);
	};

	// custom row renderer: render DetailsRow then optional subtable if expanded
	const onRenderRow = (props?: IDetailsRowProps | undefined) => {
		if (!props) return null;
		const defaultRow = <DetailsRow {...props} />;
		const item = props.item;
		const k = keyForItem(item);
		return (
			<div>
				<div className={classNames.rowWrapper}>
					{/* expand/collapse icon button (chevrons) */}
					<div className={classNames.expandButton}>
						<IconButton
							onClick={() => toggleExpand(item)}
							title={expandedKeys.indexOf(k) !== -1 ? 'Collapse details' : 'Expand details'}
							ariaLabel={expandedKeys.indexOf(k) !== -1 ? 'Collapse details' : 'Expand details'}
							iconProps={{ iconName: expandedKeys.indexOf(k) !== -1 ? 'ChevronUp' : 'ChevronDown' }}
						/>
					</div>
					<div style={{ flex: 1 }}>{defaultRow}</div>
				</div>
				{expandedKeys.indexOf(k) !== -1 && renderActionSubTable(item)}
			</div>
		);
	};

	return (
		<>
			{/* Add / export controls styled alongside RAID logs actions */}
			<div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginBottom: 8, gap: 8 }}>
				<PrimaryButton
					text="+ Add New"
					onClick={() => {
						// ensure creating mode: no selectedItem
						setSelectedItem(null);
						setIsEditing(false);
						openDialog();
					}}
				/>
				<DefaultButton
					text="Export"
					iconProps={{ iconName: 'Download' }}
					onClick={handleExportCausalAnalysis}
					disabled={!RCAItems || RCAItems.length === 0}
				/>
			</div>

			<div className={classNames.container}>
				<DetailsList
					items={paginatedItems}
					columns={displayedColumns}
					onRenderRow={onRenderRow}
					// disable selection UI and behavior
					selectionMode={SelectionMode.none}
					checkboxVisibility={CheckboxVisibility.hidden}
					compact={compact}
					className={`${className ?? ''} ${classNames.detailsList}`}
					// keep virtualization/automatic layout
					setKey="rca-table"
				/>
			</div>

			<div className={classNames.paginationBar}>
				<span style={{ fontSize: 12 }}>Page {safeCurrentPage} of {totalPages}</span>
				<div className={classNames.paginationControls}>
					<DefaultButton
						text="Previous"
						onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
						disabled={!canGoPrevious}
					/>
					<DefaultButton
						text="Next"
						onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
						disabled={!canGoNext}
					/>
				</div>
			</div>

			{isHistoryDialogOpen && (
				<div className={historyStyles.historyModalOverlay} onClick={() => setIsHistoryDialogOpen(false)}>
					<div className={historyStyles.historyModalContainer} onClick={(e) => e.stopPropagation()}>
						<div className={historyStyles.historyModalHeader}>
							<h2 className={historyStyles.historyModalTitle}>Version History{historyItemTitle ? ` - ${historyItemTitle}` : ''}</h2>
							<IconButton
								iconProps={{ iconName: 'Cancel' }}
								title="Close"
								ariaLabel="Close"
								onClick={() => setIsHistoryDialogOpen(false)}
							/>
						</div>
						<div className={historyStyles.historyModalBody}>
							<div className={historyStyles.historyContainer}>
					{isHistoryLoading ? (
						<div style={{ padding: 8 }}>Loading version history...</div>
					) : historyVersions.length === 0 ? (
						<div style={{ padding: 8 }}>No version history available.</div>
					) : (
						<div>
							<div className={historyStyles.historySummary}>Found {historyVersions.length} version(s) for this item</div>
							{historyVersions.map((version: any, index: number) => {
								const modifiedRaw = version?.Modified ?? version?.Created;
								const modifiedDate = modifiedRaw ? new Date(modifiedRaw) : undefined;
								const modifiedText = modifiedDate && !isNaN(modifiedDate.getTime()) ? modifiedDate.toLocaleString() : '-';
								const editorEmail =
									version?.Editor?.EMail ??
									version?.Editor?.Email ??
									version?.Editor?.email ??
									'-';
								const label = version?.VersionLabel ?? version?.Version ?? `${historyVersions.length - index}`;

								return (
									<div key={`${label}-${index}`} className={`${historyStyles.historyCard} ${index === 0 ? historyStyles.historyCardLatest : ''}`}>
										<div className={historyStyles.historyCardHeader}>
											<span className={historyStyles.historyVersionBadge}>Version {String(label)}</span>
											<span className={historyStyles.historyMetaText}>Modified: {modifiedText}</span>
										</div>
										<div className={historyStyles.historyMetaText}>Modified By: {String(editorEmail)}</div>
										{version?.CheckInComment ? (
											<div className={historyStyles.historyMetaText} style={{ marginTop: 2 }}>Comment: {String(version.CheckInComment)}</div>
										) : null}

										<div className={historyStyles.historySection}>
											<div className={historyStyles.historySectionTitle}>RCA Overview</div>
											<div className={historyStyles.historyFieldGrid}>
												<div className={historyStyles.historyFieldLabel}>Problem statement (Causal Analysis Trigger)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.LinkTitle)}</div>
												<div className={historyStyles.historyFieldLabel}>Cause Category</div><div className={historyStyles.historyFieldValue}>{textValue(version?.CauseCategory)}</div>
												<div className={historyStyles.historyFieldLabel}>Source</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RCASource)}</div>
												<div className={historyStyles.historyFieldLabel}>Priority</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RCAPriority)}</div>
												<div className={historyStyles.historyFieldLabel}>Related Metric (if any)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RelatedMetric)}</div>
												<div className={historyStyles.historyFieldLabel}>Related Sub Metric (if any)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RelatedSubMetric)}</div>
												<div className={historyStyles.historyFieldLabel}>Cause(s)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.Cause)}</div>
												<div className={historyStyles.historyFieldLabel}>Root Cause(s)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RootCause)}</div>
												<div className={historyStyles.historyFieldLabel}>Root Cause Analysis Technique Used and Reference (if any)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.RCATechniqueUsedAndReference)}</div>
												<div className={historyStyles.historyFieldLabel}>Type of Action</div><div className={historyStyles.historyFieldValue}>{actionTypeText(version)}</div>
											</div>
										</div>

										<div className={historyStyles.historySection}>
											<div className={historyStyles.historySectionTitle}>Correction</div>
											<div className={historyStyles.historyFieldGrid}>
												<div className={historyStyles.historyFieldLabel}>Action Plan (Correction)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.ActionPlanCorrection)}</div>
												<div className={historyStyles.historyFieldLabel}>Responsibility (Correction)</div><div className={historyStyles.historyFieldValue}>{textValue(formatResponsibilityValue(version?.ResponsibilityCorrection))}</div>
												<div className={historyStyles.historyFieldLabel}>Planned Closure Date (Correction)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.PlannedClosureDateCorrection))}</div>
												<div className={historyStyles.historyFieldLabel}>Actual Closure Date (Correction)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.ActualClosureDateCorrection))}</div>
											</div>
										</div>

										<div className={historyStyles.historySection}>
											<div className={historyStyles.historySectionTitle}>Corrective Action</div>
											<div className={historyStyles.historyFieldGrid}>
												<div className={historyStyles.historyFieldLabel}>Action Plan (Corrective Action)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.ActionPlanCorrective)}</div>
												<div className={historyStyles.historyFieldLabel}>Responsibility (Corrective Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatResponsibilityValue(version?.ResponsibilityCorrective))}</div>
												<div className={historyStyles.historyFieldLabel}>Planned Closure Date (Corrective Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.PlannedClosureDateCorrective))}</div>
												<div className={historyStyles.historyFieldLabel}>Actual Closure Date (Corrective Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.ActualClosureDateCorrective))}</div>
											</div>
										</div>

										<div className={historyStyles.historySection}>
											<div className={historyStyles.historySectionTitle}>Preventive Action And Effectiveness</div>
											<div className={historyStyles.historyFieldGrid}>
												<div className={historyStyles.historyFieldLabel}>Action Plan (Preventive Action)</div><div className={historyStyles.historyFieldValue}>{textValue(version?.ActionPlanPreventive)}</div>
												<div className={historyStyles.historyFieldLabel}>Responsibility (Preventive Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatResponsibilityValue(version?.ResponsibilityPreventive))}</div>
												<div className={historyStyles.historyFieldLabel}>Planned Closure Date (Preventive Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.PlannedClosureDatePreventive))}</div>
												<div className={historyStyles.historyFieldLabel}>Actual Closure Date (Preventive Action)</div><div className={historyStyles.historyFieldValue}>{textValue(formatDateMMDDYYYY(version?.ActualClosureDatePreventive))}</div>
												<div className={historyStyles.historyFieldLabel}>Performance before action plan</div><div className={historyStyles.historyFieldValue}>{textValue(version?.PerformanceBeforeActionPlan)}</div>
												<div className={historyStyles.historyFieldLabel}>Performance after action plan</div><div className={historyStyles.historyFieldValue}>{textValue(version?.PerformanceAfterActionPlan)}</div>
												<div className={historyStyles.historyFieldLabel}>Quantitative / Statistical effectiveness</div><div className={historyStyles.historyFieldValue}>{textValue(version?.QuantitativeOrStatisticalEffecti)}</div>
												<div className={historyStyles.historyFieldLabel}>Remarks</div><div className={historyStyles.historyFieldValue}>{textValue(version?.Remarks)}</div>
											</div>
										</div>
									</div>
								);
							})}
						</div>
					)}
						</div>
					</div>
				</div>
				</div>
			)}

			<Dialog
				hidden={!isDialogOpen}
				onDismiss={closeDialog}
				dialogContentProps={{
					type: DialogType.largeHeader,
					title: isEditing ? 'Edit RCA Item' : 'Add New RCA Item'
					
				}}
				modalProps={{
					// allow the default close (X) button to be shown
					isBlocking: true,
				}}
				minWidth={600}
				maxWidth={900}
			>
				{/* explicit close button placed top-right so it's always visible */}
				<IconButton
					iconProps={{ iconName: 'Cancel', styles: { root: { fontSize: 12 } } }}
					title="Close"
					ariaLabel="Close"
					styles={{ root: { position: 'absolute', right: 1, top: 1, zIndex: 10, width: 28, height: 28 }, icon: { fontSize: 12 } }}
					onClick={closeDialog}
				/>
				<RCAForm
					onSubmit={handleFormSubmit}
					onCancel={closeDialog}
					initialData={selectedItem ? mapRepoItemToForm(selectedItem) : {}}
					context={context}
				/>
			</Dialog>
		</>
	);
};

export default RCATable;