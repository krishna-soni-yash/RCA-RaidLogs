import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { DetailsList, DetailsRow, IDetailsRowProps, IColumn, SelectionMode, CheckboxVisibility, PrimaryButton, DefaultButton, Dialog, DialogType, IconButton, mergeStyleSets } from '@fluentui/react';
import raidStyles from '../RaidLogs/RaidLogs.module.scss';
import RCAForm from '../RootCauseAnalysisForms/RCAForm';
import { RCACOLUMNS } from '../../../../common/Constants';
import { IRCAList } from '../../../../models/IRCAList';
import { GenericService } from '../../../../services/GenericServices';
import IGenericService from '../../../../services/IGenericServices';
import { getRCAItems, RCARepository } from '../../../../repositories/RCARepository';
import { WebPartContext } from '@microsoft/sp-webpart-base';
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
			//'.ms-DetailsList': { width: '100%' },
			//'.ms-DetailsList-contentWrapper': { paddingRight: 0 },
			//'.ms-DetailsHeader': { padding: '6px 8px', background: '#faf9f8', fontSize: 12 },
			//'.ms-DetailsHeader-cell': { paddingRight: 8 },
			//'.ms-DetailsRow-cell': { padding: '6px 6px', fontSize: 13 }
		}
	},
	paginationBar: {
		display: 'flex',
		justifyContent: 'space-between',
		alignItems: 'center',
		marginTop: 12,
		padding: '0 4px'
	},
	paginationControls: {
		display: 'flex',
		columnGap: 8
	}
});



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

const formatResponsibilityValue = (input: any): string => {
	const results: string[] = [];
	const collect = (value: any): void => {
		if (value === null || value === undefined) {
			return;
		}
		if (Array.isArray(value)) {
			value.forEach(collect);
			return;
		}
		if (typeof value === 'object') {
			const email = (value as any).EMail ?? (value as any).Email ?? (value as any).email ?? (value as any).mail ?? (value as any).PrimaryEmail;
			if (email) {
				const trimmed = String(email).trim();
				if (trimmed) {
					results.push(trimmed);
				}
				return;
			}
			const title = (value as any).Title ?? (value as any).Name ?? (value as any).text;
			if (title && typeof title === 'string') {
				const sanitized = title.trim();
				if (sanitized) {
					results.push(sanitized);
				}
				return;
			}
		}
		const stringify = String(value).trim();
		if (!stringify) {
			return;
		}
		const pipeSplit = stringify.lastIndexOf('|');
		if (pipeSplit !== -1 && pipeSplit + 1 < stringify.length) {
			const afterPipe = stringify.substring(pipeSplit + 1).trim();
			if (afterPipe) {
				results.push(afterPipe);
				return;
			}
		}
		const hashSplit = stringify.lastIndexOf(';#');
		if (hashSplit !== -1 && hashSplit + 2 <= stringify.length) {
			const afterHash = stringify.substring(hashSplit + 2).trim();
			if (afterHash) {
				results.push(afterHash);
				return;
			}
		}
		results.push(stringify);
	};

	collect(input);
	const filtered = results.filter(Boolean);
	const unique: string[] = [];
	filtered.forEach((r) => {
		if (unique.indexOf(r) === -1) unique.push(r);
	});
	return unique.join(', ');
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
	const displayedColumns: IColumn[] = [
		{
			key: 'edit',
			name: '',
			fieldName: 'edit',
			minWidth: 36,
			maxWidth: 36,
			isResizable: false,
			onRender: (item: any) => (
				<IconButton
					menuIconProps={{ iconName: '' }}
					iconProps={{ iconName: 'Edit', styles: { root: { fontSize: 12 } } }}
					title="Edit"
					ariaLabel="Edit"
					styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 12 } }}
					onClick={() => {
						// open dialog with mapped initial data
						setSelectedItem(item);
						setIsEditing(true);
						setIsDialogOpen(true);
					}}
				/>
			)
		},
		...cols as IColumn[]
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
			const planned = it[`PlannedClosureDate${suffix}`] ?? '';
			const actual = it[`ActualClosureDate${suffix}`] ?? '';

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
							iconProps={{ iconName: expandedKeys.indexOf(k) !== -1 ? 'ChevronUp' : 'ChevronDown', styles: { root: { fontSize: 12 } } }}
							styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 12 } }}
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
			{/* left-aligned Add New button styled like RAID logs */}
			<div style={{ display: 'flex', justifyContent: 'flex-start', marginBottom: 8 }}>
				<PrimaryButton
					text="+ Add New"
					onClick={() => {
						// ensure creating mode: no selectedItem
						setSelectedItem(null);
						setIsEditing(false);
						openDialog();
					}}
					className={raidStyles.addButton}
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