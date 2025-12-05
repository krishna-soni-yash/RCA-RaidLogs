/*eslint-disable*/
import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILessonsLearnt, LessonsLearntDataType } from '../models/Ll Bp Rc/LessonsLearnt';
import { IBestPractices, BestPracticesDataType } from '../models/Ll Bp Rc/BestPractices';
import { IReusableComponents, IReusableComponentAttachment, ReusableComponentsDataType } from '../models/Ll Bp Rc/ReusableComponents';
import ErrorMessages from '../common/ErrorMessages';
import ILlBpRcRepository from './repositoriesInterface/Ll Bp Rc/ILlBpRcRepository';
import { SubSiteListNames } from '../common/Constants';

/**
 * Repository for Lessons Learnt, Best Practices and Reusable Components
 */
class LlBpRcrepository implements ILlBpRcRepository {
	private service: IGenericService;
	private lessonsCache: ILessonsLearnt[] | null = null;
	private bestPracticesCache: IBestPractices[] | null = null;
	private reusableCache: IReusableComponents[] | null = null;
	private lessonsCacheTimestamp = 0;
	private bestPracticesCacheTimestamp = 0;
	private reusableCacheTimestamp = 0;
	private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

	constructor(service?: IGenericService) {
		this.service = service ?? genericService;
	}



	public setService(service: IGenericService): void {
		this.service = service;
	}

	public async fetchLessonsLearnt(useCache: boolean = true, context?: WebPartContext): Promise<ILessonsLearnt[]> {
		const now = Date.now();
		if (useCache && this.lessonsCache && (now - this.lessonsCacheTimestamp) < this.CACHE_DURATION) {
			return this.lessonsCache;
		}

		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		try {
			const genericServiceInstance: IGenericService = new GenericService(undefined, context);
			genericServiceInstance.init(undefined, context);

			const items = await this.service.fetchAllItems<any>({
				context,
				listTitle: SubSiteListNames.LlBpRc,
				pageSize: 2000
			});

			const normalized = (items || [])
				.filter((it: any) => {
					const rawType = (it?.DataType ?? it?.datatype ?? '').toString().trim();
					return !rawType || rawType === LessonsLearntDataType;
				})
				.map((it: any) => ({
				ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
				LlProblemFacedLearning: it?.LlProblemFacedLearning ?? it?.ProblemFacedLearning ?? it?.Title ?? '',
				LlCategory: it?.LlCategory ?? it?.Category ?? '',
				LlSolution: it?.LlSolution ?? it?.Solution ?? '',
				LlRemarks: it?.LlRemarks ?? it?.Remarks ?? '',
				DataType: it?.DataType ?? LessonsLearntDataType
			})) as ILessonsLearnt[];

			this.lessonsCache = normalized;
			this.lessonsCacheTimestamp = now;

			return this.lessonsCache;
		} catch (error: any) {
			throw new Error('Failed to fetch Lessons Learnt: ' + (error?.message || error));
		}
	}

	public async addLessonsLearnt(item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item) {
			throw new Error('Lessons Learnt item is required.');
		}

		const payload = {
			Title: (item.LlProblemFacedLearning || item.LlSolution || '').trim() || 'Lessons Learnt',
			LlProblemFacedLearning: (item.LlProblemFacedLearning ?? '').trim(),
			LlCategory: (item.LlCategory ?? '').trim(),
			LlSolution: (item.LlSolution ?? '').trim(),
			LlRemarks: (item.LlRemarks ?? '').trim(),
			DataType: LessonsLearntDataType
		};

		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to add Lessons Learnt.');
		}

		const savedIdRaw = result.itemId ?? (result.item?.Id ?? result.item?.ID ?? result.item?.id);
		const savedId = typeof savedIdRaw === 'number' ? savedIdRaw : (savedIdRaw ? Number(savedIdRaw) : undefined);
		const hasValidId = typeof savedId === 'number' && !isNaN(savedId);

		const savedItem: ILessonsLearnt = {
			ID: hasValidId ? savedId : undefined,
			LlProblemFacedLearning: payload.LlProblemFacedLearning,
			LlCategory: payload.LlCategory,
			LlSolution: payload.LlSolution,
			LlRemarks: payload.LlRemarks,
			DataType: payload.DataType
		};

		this.invalidateLessonsCache();

		return savedItem;
	}

	public async updateLessonsLearnt(item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item?.ID) {
			throw new Error('Lessons Learnt ID is required for update.');
		}

		const payload = {
			Title: (item.LlProblemFacedLearning || item.LlSolution || '').trim() || 'Lessons Learnt',
			LlProblemFacedLearning: (item.LlProblemFacedLearning ?? '').trim(),
			LlCategory: (item.LlCategory ?? '').trim(),
			LlSolution: (item.LlSolution ?? '').trim(),
			LlRemarks: (item.LlRemarks ?? '').trim(),
			DataType: LessonsLearntDataType
		};

		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload,
			itemId: item.ID
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to update Lessons Learnt.');
		}

		this.invalidateLessonsCache();

		return {
			ID: item.ID,
			LlProblemFacedLearning: payload.LlProblemFacedLearning,
			LlCategory: payload.LlCategory,
			LlSolution: payload.LlSolution,
			LlRemarks: payload.LlRemarks,
			DataType: payload.DataType
		};
	}

	public async fetchBestPractices(useCache: boolean = true, context?: WebPartContext): Promise<IBestPractices[]> {
		const now = Date.now();
		if (useCache && this.bestPracticesCache && (now - this.bestPracticesCacheTimestamp) < this.CACHE_DURATION) {
			return this.bestPracticesCache;
		}

		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		try {
			const genericServiceInstance: IGenericService = new GenericService(undefined, context);
			genericServiceInstance.init(undefined, context);

			const items = await this.service.fetchAllItems<any>({
				context,
				listTitle: SubSiteListNames.LlBpRc,
				pageSize: 2000,
				select: [
					'ID',
					'BpBestPracticesDescription',
					'BpCategory',
					'BpReferences',
					'BpRemarks',
					'DataType'
				]
			});

			const normalized = (items || [])
				.filter((it: any) => {
					const rawType = (it?.DataType ?? it?.datatype ?? '').toString().trim();
					return !rawType || rawType === BestPracticesDataType;
				})
				.map((it: any) => {
					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						BpBestPracticesDescription: it?.BpBestPracticesDescription ?? it?.BestPracticesDescription ?? it?.Description ?? it?.Title ?? '',
						BpCategory: it?.BpCategory ?? it?.Category ?? '',
						BpReferences: it?.BpReferences ?? it?.References ?? '',
						BpRemarks: it?.BpRemarks ?? it?.Remarks ?? '',
						DataType: it?.DataType ?? BestPracticesDataType
					};
				}) as IBestPractices[];

			this.bestPracticesCache = normalized;
			this.bestPracticesCacheTimestamp = now;

			return this.bestPracticesCache;
		} catch (error: any) {
			throw new Error('Failed to fetch Best Practices: ' + (error?.message || error));
		}
	}

	public async addBestPractices(item: IBestPractices, context?: WebPartContext): Promise<IBestPractices> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item) {
			throw new Error('Best Practice item is required.');
		}

		const description = (item.BpBestPracticesDescription ?? '').trim();
		const category = (item.BpCategory ?? '').trim();

		const payload: any = {
			BpBestPracticesDescription: description,
			BpCategory: category,
			BpReferences: (item.BpReferences ?? '').trim(),
			BpRemarks: (item.BpRemarks ?? '').trim(),
			DataType: BestPracticesDataType
		};
		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to add Best Practice.');
		}

		const savedIdRaw = result.itemId ?? (result.item?.Id ?? result.item?.ID ?? result.item?.id);
		const savedId = typeof savedIdRaw === 'number' ? savedIdRaw : (savedIdRaw ? Number(savedIdRaw) : undefined);
		const hasValidId = typeof savedId === 'number' && !isNaN(savedId);

		const savedItem: IBestPractices = {
			ID: hasValidId ? savedId : undefined,
			BpBestPracticesDescription: payload.BpBestPracticesDescription,
			BpCategory: payload.BpCategory,
			BpReferences: payload.BpReferences,
			BpRemarks: payload.BpRemarks,
			DataType: payload.DataType
		};

		this.invalidateBestPracticesCache();

		return savedItem;
	}

	public async updateBestPractices(item: IBestPractices, context?: WebPartContext): Promise<IBestPractices> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item?.ID) {
			throw new Error('Best Practice ID is required for update.');
		}

		const description = (item.BpBestPracticesDescription ?? '').trim();
		const category = (item.BpCategory ?? '').trim();

		const payload: any = {
			BpBestPracticesDescription: description,
			BpCategory: category,
			BpReferences: (item.BpReferences ?? '').trim(),
			BpRemarks: (item.BpRemarks ?? '').trim(),
			DataType: BestPracticesDataType
		};
		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload,
			itemId: item.ID
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to update Best Practice.');
		}

		this.invalidateBestPracticesCache();

		return {
			ID: item.ID,
			BpBestPracticesDescription: payload.BpBestPracticesDescription,
			BpCategory: payload.BpCategory,
			BpReferences: payload.BpReferences,
			BpRemarks: payload.BpRemarks,
			DataType: payload.DataType
		};
	}

	public async fetchReusableComponents(useCache: boolean = true, context?: WebPartContext): Promise<IReusableComponents[]> {
		const now = Date.now();
		if (useCache && this.reusableCache && (now - this.reusableCacheTimestamp) < this.CACHE_DURATION) {
			return this.reusableCache;
		}

		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		try {
			const genericServiceInstance: IGenericService = new GenericService(undefined, context);
			genericServiceInstance.init(undefined, context);

			const items = await this.service.fetchAllItems<any>({
				context,
				listTitle: SubSiteListNames.LlBpRc,
				pageSize: 2000,
				select: [
					'ID',
					'RcComponentName',
					'RcLocation',
					'RcPurposeMainFunctionality',
					'RcRemarks',
					'DataType'
				]
			});

			const normalized = (items || [])
				.filter((it: any) => {
					const rawType = (it?.DataType ?? it?.datatype ?? '').toString().trim();
					return !rawType || rawType === ReusableComponentsDataType;
				})
				.map((it: any) => {
					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						RcComponentName: it?.RcComponentName ?? it?.ComponentName ?? it?.Title ?? '',
						RcLocation: it?.RcLocation ?? it?.Location ?? '',
						RcPurposeMainFunctionality: it?.RcPurposeMainFunctionality ?? it?.Purpose ?? '',
						RcRemarks: it?.RcRemarks ?? it?.Remarks ?? '',
						DataType: it?.DataType ?? ReusableComponentsDataType
					};
				}) as IReusableComponents[];

			await this.populateReusableComponentAttachments(normalized, context);

			this.reusableCache = normalized;
			this.reusableCacheTimestamp = now;

			return this.reusableCache;
		} catch (error: any) {
			throw new Error('Failed to fetch Reusable Components: ' + (error?.message || error));
		}
	}

	public async addReusableComponents(item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item) {
			throw new Error('Reusable Component item is required.');
		}

		const payload: any = {
			RcComponentName: (item.RcComponentName ?? '').trim(),
			RcLocation: (item.RcLocation ?? '').trim(),
			RcPurposeMainFunctionality: (item.RcPurposeMainFunctionality ?? '').trim(),
			RcRemarks: (item.RcRemarks ?? '').trim(),
			DataType: ReusableComponentsDataType
		};
		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to add Reusable Component.');
		}

		const savedIdRaw = result.itemId ?? (result.item?.Id ?? result.item?.ID ?? result.item?.id);
		const savedId = typeof savedIdRaw === 'number' ? savedIdRaw : (savedIdRaw ? Number(savedIdRaw) : undefined);
		const hasValidId = typeof savedId === 'number' && !isNaN(savedId);

		const savedItem: IReusableComponents = {
			ID: hasValidId ? savedId : undefined,
			RcComponentName: payload.RcComponentName,
			RcLocation: payload.RcLocation,
			RcPurposeMainFunctionality: payload.RcPurposeMainFunctionality,
			RcRemarks: payload.RcRemarks,
			DataType: payload.DataType
		};

		if (hasValidId && context) {
			const filesToUpload = this.extractNewAttachmentFiles(item);
			if (filesToUpload.length > 0) {
				await this.uploadReusableComponentAttachments(savedId as number, filesToUpload, context);
			}
			savedItem.attachments = await this.getReusableComponentAttachments(savedId as number, context);
		} else {
			savedItem.attachments = [];
		}

		savedItem.newAttachments = [];

		this.invalidateReusableCache();

		return savedItem;
	}

	public async updateReusableComponents(item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents> {
		if (!context) {
			throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
		}

		if (!item?.ID) {
			throw new Error('Reusable Component ID is required for update.');
		}

		const payload: any = {
			RcComponentName: (item.RcComponentName ?? '').trim(),
			RcLocation: (item.RcLocation ?? '').trim(),
			RcPurposeMainFunctionality: (item.RcPurposeMainFunctionality ?? '').trim(),
			RcRemarks: (item.RcRemarks ?? '').trim(),
			DataType: ReusableComponentsDataType
		};
		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload,
			itemId: item.ID
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to update Reusable Component.');
		}

		const filesToUpload = this.extractNewAttachmentFiles(item);
		if (filesToUpload.length > 0) {
			await this.uploadReusableComponentAttachments(item.ID, filesToUpload, context);
		}

		const attachments = await this.getReusableComponentAttachments(item.ID, context);

		this.invalidateReusableCache();

		return {
			ID: item.ID,
			RcComponentName: payload.RcComponentName,
			RcLocation: payload.RcLocation,
			RcPurposeMainFunctionality: payload.RcPurposeMainFunctionality,
			RcRemarks: payload.RcRemarks,
			DataType: payload.DataType,
			attachments,
			newAttachments: []
		};
	}

	private isFileInstance(value: unknown): value is File {
		return typeof File !== 'undefined' && value instanceof File;
	}

	private extractNewAttachmentFiles(item?: IReusableComponents | null): File[] {
		if (!item) {
			return [];
		}

		const files: File[] = [];

		if (Array.isArray(item.newAttachments)) {
			for (const maybeFile of item.newAttachments) {
				if (this.isFileInstance(maybeFile)) {
					files.push(maybeFile);
				}
			}
		}

		if (Array.isArray(item.attachments)) {
			for (const maybeFile of item.attachments as unknown[]) {
				if (this.isFileInstance(maybeFile)) {
					files.push(maybeFile);
				}
			}
		}

		return files;
	}

	private async getReusableComponentAttachments(itemId: number, context: WebPartContext): Promise<IReusableComponentAttachment[]> {
		if (!context || !itemId) {
			return [];
		}

		try {
			const targetSiteUrl = this.service.getSiteUrlForList(SubSiteListNames.LlBpRc, context);
			const sp = await this.service.getSpInstanceForSite(targetSiteUrl, context);
			const files = await sp?.web?.lists?.getByTitle?.(SubSiteListNames.LlBpRc)?.items?.getById?.(itemId)?.attachmentFiles?.();
			if (!files) {
				return [];
			}

			return (files as any[]).map((file: any) => ({
				FileName: file?.FileName ?? file?.FileLeafRef ?? '',
				ServerRelativeUrl: file?.ServerRelativeUrl ?? file?.ServerRelativePath?.DecodedUrl ?? ''
			})) as IReusableComponentAttachment[];
		} catch (error) {
			console.warn('LlBpRcrepository.getReusableComponentAttachments: failed to load attachments', { itemId, error });
			return [];
		}
	}

	private async uploadReusableComponentAttachments(itemId: number, files: File[], context: WebPartContext): Promise<void> {
		if (!context || !itemId || !Array.isArray(files) || files.length === 0) {
			return;
		}

		try {
			const targetSiteUrl = this.service.getSiteUrlForList(SubSiteListNames.LlBpRc, context);
			const sp = await this.service.getSpInstanceForSite(targetSiteUrl, context);
			const listRef = sp?.web?.lists?.getByTitle?.(SubSiteListNames.LlBpRc);
			const itemRef = listRef?.items?.getById?.(itemId);

			if (!itemRef || typeof itemRef.attachmentFiles?.add !== 'function') {
				throw new Error('Attachment API not available for Reusable Components list.');
			}

			for (const file of files) {
				if (!this.isFileInstance(file)) {
					continue;
				}
				await itemRef.attachmentFiles.add(file.name, file);
			}
		} catch (error) {
			console.error('LlBpRcrepository.uploadReusableComponentAttachments: failed to upload attachment(s)', { itemId, error });
			throw error;
		}
	}

	private async populateReusableComponentAttachments(items: IReusableComponents[], context: WebPartContext): Promise<void> {
		if (!Array.isArray(items) || items.length === 0) {
			return;
		}

		const targetSiteUrl = this.service.getSiteUrlForList(SubSiteListNames.LlBpRc, context);
		const sp = await this.service.getSpInstanceForSite(targetSiteUrl, context);
		const listRef = sp?.web?.lists?.getByTitle?.(SubSiteListNames.LlBpRc);

		await Promise.all(items.map(async (component) => {
			if (!component?.ID) {
				component.attachments = [];
				component.newAttachments = [];
				return;
			}

			if (!listRef?.items?.getById) {
				component.attachments = [];
				component.newAttachments = [];
				return;
			}

			try {
				const files = await listRef.items.getById(component.ID).attachmentFiles();
				component.attachments = (files || []).map((file: any) => ({
					FileName: file?.FileName ?? file?.FileLeafRef ?? '',
					ServerRelativeUrl: file?.ServerRelativeUrl ?? file?.ServerRelativePath?.DecodedUrl ?? ''
				})) as IReusableComponentAttachment[];
				component.newAttachments = [];
			} catch (error) {
				console.warn('LlBpRcrepository.populateReusableComponentAttachments: failed for item', component.ID, error);
				component.attachments = [];
				component.newAttachments = [];
			}
		}));
	}

	private invalidateLessonsCache(): void {
		this.lessonsCache = null;
		this.lessonsCacheTimestamp = 0;
	}

	private invalidateBestPracticesCache(): void {
		this.bestPracticesCache = null;
		this.bestPracticesCacheTimestamp = 0;
	}

	private invalidateReusableCache(): void {
		this.reusableCache = null;
		this.reusableCacheTimestamp = 0;
	}

	public refresh(): void {
		this.lessonsCache = null;
		this.bestPracticesCache = null;
		this.reusableCache = null;

		this.lessonsCacheTimestamp = 0;
		this.bestPracticesCacheTimestamp = 0;
		this.reusableCacheTimestamp = 0;
	}
}

const defaultInstance = new LlBpRcrepository();

export default defaultInstance;
export const LlBpRcRepo = defaultInstance;

export const fetchLessonsLearnt = async (useCache: boolean = false, context?: WebPartContext): Promise<ILessonsLearnt[]> => defaultInstance.fetchLessonsLearnt(useCache, context);
export const addLessonsLearnt = async (item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt> => defaultInstance.addLessonsLearnt(item, context);
export const updateLessonsLearnt = async (item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt> => defaultInstance.updateLessonsLearnt(item, context);

export const fetchBestPractices = async (useCache: boolean = false, context?: WebPartContext): Promise<IBestPractices[]> => defaultInstance.fetchBestPractices(useCache, context);
export const fetchReusableComponents = async (useCache: boolean = false, context?: WebPartContext): Promise<IReusableComponents[]> => defaultInstance.fetchReusableComponents(useCache, context);
export const addBestPractices = async (item: IBestPractices, context?: WebPartContext): Promise<IBestPractices> => defaultInstance.addBestPractices(item, context);
export const updateBestPractices = async (item: IBestPractices, context?: WebPartContext): Promise<IBestPractices> => defaultInstance.updateBestPractices(item, context);
export const addReusableComponents = async (item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents> => defaultInstance.addReusableComponents(item, context);
export const updateReusableComponents = async (item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents> => defaultInstance.updateReusableComponents(item, context);