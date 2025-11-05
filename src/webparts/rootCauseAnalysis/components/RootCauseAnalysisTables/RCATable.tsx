import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, IColumn, SelectionMode, CheckboxVisibility, DefaultButton, Dialog, DialogType } from '@fluentui/react';
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

	useEffect(() => {
		fetchRCAItems();
	}, []);
	; const fetchRCAItems = async () => {
		const genericServiceInstance: IGenericService = new GenericService(undefined, context);
		genericServiceInstance.init(undefined, context);
		const RCARepo: IRCARepository = new RCARepository(genericServiceInstance);
		RCARepo.setService(genericServiceInstance);
		const RAitems = await getRCAItems(true, context);
		setRCAItems(RAitems);
	}
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
				items={RCAItems.length > 0 ? RCAItems : localItems}
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