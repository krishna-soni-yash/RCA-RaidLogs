import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, IColumn, SelectionMode, CheckboxVisibility, DefaultButton, Dialog, DialogType } from '@fluentui/react';
import RCAForm from '../RootCauseAnalysisForms/RCAForm';

// replace dynamic column creation with fixed RCA columns
const RCACOLUMNS: IColumn[] = [
	{
		key: 'problemStatement',
		name: 'Problem statement (Causal Analysis Trigger)',
		fieldName: 'problemStatement',
		minWidth: 150,
		maxWidth: 350,
		isResizable: true
	},
	{ key: 'causeCategory', name: 'Cause Category', fieldName: 'causeCategory', minWidth: 100, maxWidth: 200, isResizable: true },
	{ key: 'source', name: 'Source', fieldName: 'source', minWidth: 100, maxWidth: 200, isResizable: true },
	{ key: 'priority', name: 'Priority', fieldName: 'priority', minWidth: 80, maxWidth: 120, isResizable: true },
	{ key: 'relatedMetric', name: 'Related Metric (if any)', fieldName: 'relatedMetric', minWidth: 140, maxWidth: 250, isResizable: true },
	{ key: 'causes', name: 'Cause(s)', fieldName: 'causes', minWidth: 150, maxWidth: 300, isResizable: true },
	{ key: 'rootCauses', name: 'Root Cause(s)', fieldName: 'rootCauses', minWidth: 150, maxWidth: 300, isResizable: true },
	{ key: 'analysisTechnique', name: 'Root Cause Analysis Technique Used and Reference (if any)', fieldName: 'analysisTechnique', minWidth: 180, maxWidth: 350, isResizable: true },
	{ key: 'actionType', name: 'Type of Action', fieldName: 'actionType', minWidth: 120, maxWidth: 200, isResizable: true },
	{ key: 'actionPlan', name: 'Action Plan', fieldName: 'actionPlan', minWidth: 150, maxWidth: 350, isResizable: true },
	{ key: 'responsibility', name: 'Responsibility', fieldName: 'responsibility', minWidth: 120, maxWidth: 200, isResizable: true },
	{ key: 'plannedClosureDate', name: 'Planned Closure Date', fieldName: 'plannedClosureDate', minWidth: 120, maxWidth: 150, isResizable: true },
	{ key: 'actualClosureDate', name: 'Actual Closure Date', fieldName: 'actualClosureDate', minWidth: 120, maxWidth: 150, isResizable: true },
	{ key: 'performanceBefore', name: 'Performance before action plan', fieldName: 'performanceBefore', minWidth: 150, maxWidth: 220, isResizable: true },
	{ key: 'performanceAfter', name: 'Performance after action plan', fieldName: 'performanceAfter', minWidth: 150, maxWidth: 220, isResizable: true },
	{ key: 'quantitativeEffectiveness', name: 'Quantitative / Statistical effectiveness', fieldName: 'quantitativeEffectiveness', minWidth: 180, maxWidth: 260, isResizable: true },
	{ key: 'remarks', name: 'Remarks', fieldName: 'remarks', minWidth: 120, maxWidth: 300, isResizable: true }
];

interface RCATableProps {
	// rows to show in the table
	items: any[];
	// optional preset columns; if omitted columns will be generated from the first item
	columns?: IColumn[];
	// optional compact mode
	compact?: boolean;
	// optional className for styling
	className?: string;
}

const createColumnsFromItems = (items: any[]): IColumn[] => {
	// keep fallback for arbitrary items, but prefer RCACOLUMNS as default
	if (!items || items.length === 0) {
		return RCACOLUMNS;
	}
	const first = items[0];
	// if items already match RCA structure, return RCACOLUMNS to preserve header names
	const hasRCAKeys = Object.prototype.hasOwnProperty.call(first, 'problemStatement');
	return hasRCAKeys ? RCACOLUMNS : Object.keys(first).map((key) => ({
		key,
		name: key,
		fieldName: key,
		minWidth: 50,
		maxWidth: 300,
		isResizable: true,
	}));
};

const RCATable: React.FC<RCATableProps> = ({ items, columns, compact = false, className }) => {
	// prefer passed columns, then RCACOLUMNS, then fallback dynamic columns
	const cols = columns && columns.length ? columns : (items && items.length ? createColumnsFromItems(items) : RCACOLUMNS);

	// local state so we can append new items added via the form dialog
	const [localItems, setLocalItems] = useState<any[]>(items || []);
	const [isDialogOpen, setIsDialogOpen] = useState(false);

	useEffect(() => {
		setLocalItems(items || []);
	}, [items]);

	const openDialog = () => setIsDialogOpen(true);
	const closeDialog = () => setIsDialogOpen(false);

	const handleFormSubmit = (data: any) => {
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

	return (
		<>
			{/* right-aligned button at the top */}
			<div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
				<DefaultButton
					text="Add New Item"
					onClick={openDialog}
				/>
			</div>

			<DetailsList
				items={localItems}
				columns={cols}
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
					title: 'Add New RCA Item'
				}}
				minWidth={600}
				maxWidth={900}
			>
				<RCAForm onSubmit={handleFormSubmit} initialData={{}} />
			</Dialog>
		</>
	);
};

export default RCATable;