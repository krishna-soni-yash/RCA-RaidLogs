/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFetchOptions {
	context?: WebPartContext;
	spInstance?: any;
	listTitle: string;
	select?: string[];
	expand?: string[];
	filter?: string;
	orderBy?: string;
	pageSize?: number;
	maxRetries?: number;
	retryDelayMs?: number;
	forceSiteUrl?: string;
}

export interface ISaveOptions {
	context?: WebPartContext;
	spInstance?: any;
	listTitle: string;
	item: any;
	itemId?: number;
	select?: string[];
	expand?: string[];
	maxRetries?: number;
	retryDelayMs?: number;
	forceSiteUrl?: string;
	validateColumns?: boolean;
}

export interface ISaveResult<T = any> {
	success: boolean;
	item?: T;
	itemId?: number;
	error?: string;
}



export interface IBatchSaveOptions {
	context?: WebPartContext;
	spInstance?: any;
	listTitle: string;
	items: any[];
	batchSize?: number;
	maxRetries?: number;
	retryDelayMs?: number;
	forceSiteUrl?: string;
	validateColumns?: boolean;
}

export interface IUpdateOptions {
	context?: WebPartContext;
	spInstance?: any;
	listTitle: string;
	itemId: number;
	item: any;
	select?: string[];
	expand?: string[];
	maxRetries?: number;
	retryDelayMs?: number;
	forceSiteUrl?: string;
	validateColumns?: boolean;
}

export interface IGenericService {
	// Initialize or replace the sp instance
	init(spInstance?: any, context?: WebPartContext): void;

	// Fetch all items from a list with options
	fetchAllItems<T = any>(options: IFetchOptions): Promise<T[]>;

	// Save a single item to a list (create or update)
	saveItem<T = any>(options: ISaveOptions): Promise<ISaveResult<T>>;

	// Save multiple items to a list in batches
	saveBatchItems<T = any>(options: IBatchSaveOptions): Promise<ISaveResult<T>[]>;

	// Update an existing item in a list
	updateItem<T = any>(options: IUpdateOptions): Promise<ISaveResult<T>>;

	// Clear cached SP instances
	clearCache(): void;
}

export default IGenericService;

