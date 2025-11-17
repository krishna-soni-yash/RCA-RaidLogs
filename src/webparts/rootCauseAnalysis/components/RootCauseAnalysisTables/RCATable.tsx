import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, IColumn, SelectionMode, CheckboxVisibility, DefaultButton, Dialog, DialogType, IconButton } from '@fluentui/react';
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
			)
		},
		...cols as IColumn[]
	];

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
					iconProps={{ iconName: 'Cancel' }}
					title="Close"
					ariaLabel="Close"
					onClick={closeDialog}
					style={{
						position: 'absolute',
						right: 1,
						top: 1,
						zIndex: 10
					}}
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