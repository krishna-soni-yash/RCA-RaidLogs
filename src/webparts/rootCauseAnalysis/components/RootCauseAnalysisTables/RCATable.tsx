import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, DetailsRow, IDetailsRowProps, IColumn, SelectionMode, CheckboxVisibility, DefaultButton, Dialog, DialogType, IconButton } from '@fluentui/react';
import RCAForm from '../RootCauseAnalysisForms/RCAForm';
import { RCACOLUMNS } from '../../../../common/Constants';
import { IRCAList } from '../../../../models/IRCAList';
import { GenericService } from '../../../../services/GenericServices';
import IGenericService from '../../../../services/IGenericServices';
import { getRCAItems, RCARepository } from '../../../../repositories/repositoriesInterface/RCARepository';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IRCARepository from '../../../../repositories/IRCARepository';



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
		fetchRCAItems();
	}, [context]);
	; const fetchRCAItems = async () => {
		const genericServiceInstance: IGenericService = new GenericService(undefined, context);
		genericServiceInstance.init(undefined, context);
		const RCARepo: IRCARepository = new RCARepository(genericServiceInstance);
		RCARepo.setService(genericServiceInstance);
		const RAitems = await getRCAItems(true, context);
		setRCAItems(RAitems);
	}
	// helper: map repository item (IRCAList) to RCAForm initialData shape
	const mapRepoItemToForm = (it: any): any => {
		if (!it) return {};
		const form: any = {};
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
				// Responsibility fields from repo are strings like email; keep as-is or as array
				responsibility: it[`Responsibility${suffix}`] || it[`Responsibility${suffix}`] || '',
				plannedClosureDate: it[`PlannedClosureDate${suffix}`] ? new Date(it[`PlannedClosureDate${suffix}`]) : undefined,
				actualClosureDate: it[`ActualClosureDate${suffix}`] ? new Date(it[`ActualClosureDate${suffix}`]) : undefined
			};
		});

		form.actionDetails = details;
		form.performanceBefore = it.PerformanceBeforeActionPlan || '';
		form.performanceAfter = it.PerformanceAfterActionPlan || '';
		form.quantitativeEffectiveness = it.QuantitativeOrStatisticalEffecti || '';
		form.remarks = it.Remarks || '';
		// preserve id for editing context
		form.__repoId = it.ID ?? it.Id ?? it.Id;
		return form;
	};

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
	const [expandedInitialized, setExpandedInitialized] = useState<boolean>(false);
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

	// initialize expanded state once when items (remote or local) become available
	useEffect(() => {
		if (expandedInitialized) return;
		const items = (RCAItems && RCAItems.length > 0) ? RCAItems : (localItems && localItems.length > 0 ? localItems : []);
		if (items.length === 0) return;
		const keys = items.map((it: any) => keyForItem(it));
		setExpandedKeys(keys);
		setExpandedInitialized(true);
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [RCAItems, localItems]);

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
			const responsibility = it[`Responsibility${suffix}`] ?? '';
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
			{ key: 'type', name: 'Type of Action', fieldName: 'type', minWidth: 120, maxWidth: 180, isResizable: true },
			{ key: 'actionPlan', name: 'Action Plan', fieldName: 'actionPlan', minWidth: 250, maxWidth: 500, isResizable: true },
			{ key: 'responsibility', name: 'Responsibility', fieldName: 'responsibility', minWidth: 180, maxWidth: 300, isResizable: true },
			{ key: 'planned', name: 'Planned', fieldName: 'planned', minWidth: 120, maxWidth: 160, isResizable: true },
			{ key: 'actual', name: 'Actual', fieldName: 'actual', minWidth: 120, maxWidth: 160, isResizable: true }
		];

		return (
			<div style={{ padding: '8px 12px 12px 56px', background: '#fafafa', borderLeft: '2px solid #e1dfdd' }}>
				<DetailsList
					items={rows}
					columns={subColumns}
					selectionMode={SelectionMode.none}
					checkboxVisibility={CheckboxVisibility.hidden}
					compact={true}
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
				<div style={{ display: 'flex', alignItems: 'center' }}>
					{/* expand/collapse icon button (chevrons) */}
					<IconButton
						onClick={() => toggleExpand(item)}
						title={expandedKeys.indexOf(k) !== -1 ? 'Collapse details' : 'Expand details'}
						ariaLabel={expandedKeys.indexOf(k) !== -1 ? 'Collapse details' : 'Expand details'}
						iconProps={{ iconName: expandedKeys.indexOf(k) !== -1 ? 'ChevronUp' : 'ChevronDown', styles: { root: { fontSize: 12 } } }}
						styles={{ root: { width: 28, height: 28, marginLeft: 8, marginRight: 8 }, icon: { fontSize: 12 } }}
					/>
					<div style={{ flex: 1 }}>{defaultRow}</div>
				</div>
				{expandedKeys.indexOf(k) !== -1 && renderActionSubTable(item)}
			</div>
		);
	};

	return (
		<>
			{/* right-aligned button at the top */}
			<div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
				<DefaultButton
					text="Add New Item"
					onClick={() => {
						// ensure creating mode: no selectedItem
						setSelectedItem(null);
						setIsEditing(false);
						openDialog();
					}}
				/>
			</div>

			<DetailsList
				items={RCAItems.length > 0 ? RCAItems : localItems}
				columns={displayedColumns}
				onRenderRow={onRenderRow}
				// disable selection UI and behavior
				selectionMode={SelectionMode.none}
				checkboxVisibility={CheckboxVisibility.hidden}
				compact={compact}
				className={className}
				// keep virtualization/automatic layout
				setKey="rca-table"
			/>

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
 					initialData={selectedItem ? mapRepoItemToForm(selectedItem) : {}}
 					context={context}
 				/>
			</Dialog>
		</>
	);
};

export default RCATable;