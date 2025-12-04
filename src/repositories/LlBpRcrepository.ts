import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILessonsLearnt, LessonsLearntDataType } from '../models/Ll Bp Rc/LessonsLearnt';
import { IBestPractices, BestPracticesDataType } from '../models/Ll Bp Rc/BestPractices';
import { IReusableComponents, ReusableComponentsDataType } from '../models/Ll Bp Rc/ReusableComponents';
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

	private toArray<T>(value?: T | T[]): T[] {
		if (value === undefined || value === null) {
			return [];
		}
		return Array.isArray(value) ? value : [value];
	}

	private pickFirstString(values: Array<string | undefined | null>): string | undefined {
		for (const value of values) {
			if (typeof value === 'string') {
				const trimmed = value.trim();
				if (trimmed.length > 0) {
					return trimmed;
				}
			}
		}
		return undefined;
	}

	private pickFirstNumber(values: any[]): number | undefined {
		for (const value of values) {
			if (typeof value === 'number' && !isNaN(value)) {
				return value;
			}
			if (typeof value === 'string' && value.trim().length > 0) {
				const parsed = Number(value.trim());
				if (!isNaN(parsed)) {
					return parsed;
				}
			}
		}
		return undefined;
	}

	private normalizeIdArray(value: number | number[] | string | string[] | undefined | null): number[] {
		const results: number[] = [];
		if (value === undefined || value === null) {
			return results;
		}
		const entries = this.toArray<any>(value as any);
		for (const entry of entries) {
			if (typeof entry === 'number' && !isNaN(entry)) {
				results.push(entry);
			} else if (typeof entry === 'string') {
				const parsed = Number(entry.trim());
				if (!isNaN(parsed)) {
					results.push(parsed);
				}
			}
		}
		return results;
	}

	private assignMultiLookupId(
		target: Record<string, any>,
		fieldName: string,
		resolvedId?: number,
		originalValue?: number | number[] | string | string[] | null,
		forceClear: boolean = false
	): number[] | undefined {
		const ids: number[] = [];
		if (typeof resolvedId === 'number' && !isNaN(resolvedId)) {
			ids.push(resolvedId);
		} else {
			ids.push(...this.normalizeIdArray(originalValue));
		}
		const distinctIds: number[] = [];
		for (const id of ids) {
			if (typeof id === 'number' && !isNaN(id) && distinctIds.indexOf(id) === -1) {
				distinctIds.push(id);
			}
		}
		if (distinctIds.length > 0) {
			target[fieldName] = { results: distinctIds };
			return distinctIds;
		}
		if (forceClear) {
			target[fieldName] = { results: [] };
			return [];
		}
		return undefined;
	}

	private normalizePersonField(item: any, fieldName: string): { displayName?: string; email?: string; loginName?: string; id?: number } {
		if (!item) {
			return {};
		}
		const rawValue = item[fieldName];
		const personObject = Array.isArray(rawValue) ? (rawValue.length > 0 ? rawValue[0] : undefined) : rawValue;
		const displayName = typeof rawValue === 'string'
			? rawValue
			: this.pickFirstString([
				personObject?.Title,
				personObject?.Name,
				personObject?.DisplayName,
				personObject?.LoginName,
				personObject?.UserPrincipalName,
				personObject?.Email,
				personObject?.EMail
			]);
		const email = this.pickFirstString([
			item?.[`${fieldName}Email`],
			personObject?.Email,
			personObject?.EMail
		]);
		const loginName = this.pickFirstString([
			item?.[`${fieldName}LoginName`],
			personObject?.LoginName,
			personObject?.UserPrincipalName,
			personObject?.Name
		]);
		const id = this.pickFirstNumber([
			item?.[`${fieldName}Id`],
			personObject?.Id,
			personObject?.ID
		]);
		return {
			displayName,
			email,
			loginName,
			id
		};
	}

	private async resolvePersonForSave(
		context: WebPartContext,
		person: {
			id?: number | number[];
			loginName?: string | string[];
			email?: string | string[];
			displayName?: string | string[];
		}
	): Promise<{ id?: number; loginName?: string; email?: string; displayName?: string }> {
		const existingId = this.pickFirstNumber(this.toArray(person?.id));
		const loginCandidate = this.pickFirstString(this.toArray(person?.loginName));
		const emailCandidate = this.pickFirstString(this.toArray(person?.email));
		const displayCandidate = this.pickFirstString([
			...this.toArray(person?.displayName),
			loginCandidate,
			emailCandidate
		]);
		if (existingId !== undefined) {
			return {
				id: existingId,
				loginName: loginCandidate,
				email: emailCandidate,
				displayName: displayCandidate
			};
		}

		const ensureCandidates = [loginCandidate, emailCandidate]
			.filter((candidate, index, array) => typeof candidate === 'string' && candidate.length > 0 && array.indexOf(candidate) === index) as string[];

		for (const candidate of ensureCandidates) {
			try {
				const resolver: any = typeof (this.service as any)?.ensureuser === 'function'
					? this.service
					: new GenericService(undefined, context);
				const ensured = await resolver.ensureuser(context, candidate);
				if (ensured) {
					const ensuredId = this.pickFirstNumber([
						ensured?.Id,
						ensured?.ID,
						ensured?.UserId
					]);
					if (ensuredId !== undefined) {
						return {
							id: ensuredId,
							loginName: this.pickFirstString([
								ensured?.LoginName,
								ensured?.UserPrincipalName,
								loginCandidate,
								emailCandidate
							]),
							email: this.pickFirstString([
								ensured?.Email,
								ensured?.EMail,
								emailCandidate,
								loginCandidate
							]),
							displayName: this.pickFirstString([
								ensured?.Title,
								ensured?.DisplayName,
								ensured?.LoginName,
								emailCandidate,
								loginCandidate,
								displayCandidate
							])
						};
					}
				}
			} catch (error) {
				console.warn('LlBpRcrepository.resolvePersonForSave: failed to ensure user', candidate, error);
			}
		}

		return {
			id: undefined,
			loginName: loginCandidate,
			email: emailCandidate,
			displayName: displayCandidate
		};
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
					'BpReferences',
					'BpResponsibility',
					'BpResponsibilityId',
					'BpRemarks',
					'DataType',
					'BpResponsibility/Id',
					'BpResponsibility/Title',
					'BpResponsibility/EMail'
				],
				expand: ['BpResponsibility']
			});

			const normalized = (items || [])
				.filter((it: any) => {
					const rawType = (it?.DataType ?? it?.datatype ?? '').toString().trim();
					return !rawType || rawType === BestPracticesDataType;
				})
				.map((it: any) => {
					const responsibilityInfo = this.normalizePersonField(it, 'BpResponsibility');
					const responsibilityDisplay = this.pickFirstString([
						typeof it?.BpResponsibility === 'string' ? it.BpResponsibility : undefined,
						responsibilityInfo.displayName
					]) ?? '';

					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						BpBestPracticesDescription: it?.BpBestPracticesDescription ?? it?.BestPracticesDescription ?? it?.Description ?? it?.Title ?? '',
						BpReferences: it?.BpReferences ?? it?.References ?? '',
						BpResponsibility: responsibilityDisplay,
						BpResponsibilityId: responsibilityInfo.id,
						BpResponsibilityEmail: responsibilityInfo.email,
						BpResponsibilityLoginName: responsibilityInfo.loginName,
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
		const resolvedResponsibility = await this.resolvePersonForSave(context, {
			id: item.BpResponsibilityId,
			loginName: item.BpResponsibilityLoginName,
			email: item.BpResponsibilityEmail,
			displayName: item.BpResponsibility
		});
		
		const payload: any = {
			BpBestPracticesDescription: description,
			BpReferences: (item.BpReferences ?? '').trim(),
			BpRemarks: (item.BpRemarks ?? '').trim(),
			DataType: BestPracticesDataType
		};
		const responsibilityIds = this.assignMultiLookupId(payload, 'BpResponsibilityId', resolvedResponsibility.id, item.BpResponsibilityId);
		payload.BpResponsibilityId = responsibilityIds;
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
			BpReferences: payload.BpReferences,
			BpResponsibility: resolvedResponsibility.displayName,
			BpResponsibilityId: responsibilityIds ? (responsibilityIds.length === 1 ? responsibilityIds[0] : responsibilityIds) : undefined,
			BpResponsibilityEmail: resolvedResponsibility.email,
			BpResponsibilityLoginName: resolvedResponsibility.loginName,
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
		const resolvedResponsibility = await this.resolvePersonForSave(context, {
			id: item.BpResponsibilityId,
			loginName: item.BpResponsibilityLoginName,
			email: item.BpResponsibilityEmail,
			displayName: item.BpResponsibility
		});
		
		const payload: any = {
			BpBestPracticesDescription: description,
			BpReferences: (item.BpReferences ?? '').trim(),
			BpRemarks: (item.BpRemarks ?? '').trim(),
			DataType: BestPracticesDataType
		};
		const responsibilityIds = this.assignMultiLookupId(payload, 'BpResponsibilityId', resolvedResponsibility.id, item.BpResponsibilityId, true);
		payload.BpResponsibilityId = responsibilityIds;
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
			BpReferences: payload.BpReferences,
			BpResponsibility: resolvedResponsibility.displayName,
			BpResponsibilityId: responsibilityIds ? (responsibilityIds.length === 1 ? responsibilityIds[0] : responsibilityIds) : undefined,
			BpResponsibilityEmail: resolvedResponsibility.email,
			BpResponsibilityLoginName: resolvedResponsibility.loginName,
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
					'RcResponsibility',
					'RcResponsibilityId',
					'RcRemarks',
					'DataType',
					'RcResponsibility/Id',
					'RcResponsibility/Title',
					'RcResponsibility/EMail'
				],
				expand: ['RcResponsibility']
			});

			const normalized = (items || [])
				.filter((it: any) => {
					const rawType = (it?.DataType ?? it?.datatype ?? '').toString().trim();
					return !rawType || rawType === ReusableComponentsDataType;
				})
				.map((it: any) => {
					const responsibilityInfo = this.normalizePersonField(it, 'RcResponsibility');
					const responsibilityDisplay = this.pickFirstString([
						typeof it?.RcResponsibility === 'string' ? it.RcResponsibility : undefined,
						responsibilityInfo.displayName
					]) ?? '';

					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						RcComponentName: it?.RcComponentName ?? it?.ComponentName ?? it?.Title ?? '',
						RcLocation: it?.RcLocation ?? it?.Location ?? '',
						RcPurposeMainFunctionality: it?.RcPurposeMainFunctionality ?? it?.Purpose ?? '',
						RcResponsibility: responsibilityDisplay,
						RcResponsibilityId: responsibilityInfo.id,
						RcResponsibilityEmail: responsibilityInfo.email,
						RcResponsibilityLoginName: responsibilityInfo.loginName,
						RcRemarks: it?.RcRemarks ?? it?.Remarks ?? '',
						DataType: it?.DataType ?? ReusableComponentsDataType
					};
				}) as IReusableComponents[];

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

		const resolvedResponsibility = await this.resolvePersonForSave(context, {
			id: item.RcResponsibilityId,
			loginName: item.RcResponsibilityLoginName,
			email: item.RcResponsibilityEmail,
			displayName: item.RcResponsibility
		});

		const payload: any = {
			RcComponentName: (item.RcComponentName ?? '').trim(),
			RcLocation: (item.RcLocation ?? '').trim(),
			RcPurposeMainFunctionality: (item.RcPurposeMainFunctionality ?? '').trim(),
			RcRemarks: (item.RcRemarks ?? '').trim(),
			DataType: ReusableComponentsDataType
		};
		const responsibilityIds = this.assignMultiLookupId(payload, 'RcResponsibilityId', resolvedResponsibility.id, item.RcResponsibilityId);
		payload.RcResponsibilityId = responsibilityIds;
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
			RcResponsibility: payload.RcResponsibility,
			RcResponsibilityId: responsibilityIds ? (responsibilityIds.length === 1 ? responsibilityIds[0] : responsibilityIds) : undefined,
			RcResponsibilityEmail: resolvedResponsibility.email,
			RcResponsibilityLoginName: resolvedResponsibility.loginName,
			RcRemarks: payload.RcRemarks,
			DataType: payload.DataType
		};

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

		const resolvedResponsibility = await this.resolvePersonForSave(context, {
			id: item.RcResponsibilityId,
			loginName: item.RcResponsibilityLoginName,
			email: item.RcResponsibilityEmail,
			displayName: item.RcResponsibility
		});

		const payload: any = {
			RcComponentName: (item.RcComponentName ?? '').trim(),
			RcLocation: (item.RcLocation ?? '').trim(),
			RcPurposeMainFunctionality: (item.RcPurposeMainFunctionality ?? '').trim(),
			RcRemarks: (item.RcRemarks ?? '').trim(),
			DataType: ReusableComponentsDataType
		};
		const responsibilityIds = this.assignMultiLookupId(payload, 'RcResponsibilityId', resolvedResponsibility.id, item.RcResponsibilityId, true);
		payload.RcResponsibilityId = responsibilityIds;
		const result = await this.service.saveItem<any>({
			context,
			listTitle: SubSiteListNames.LlBpRc,
			item: payload,
			itemId: item.ID
		});

		if (!result?.success) {
			throw new Error(result?.error ?? 'Failed to update Reusable Component.');
		}

		this.invalidateReusableCache();

		return {
			ID: item.ID,
			RcComponentName: payload.RcComponentName,
			RcLocation: payload.RcLocation,
			RcPurposeMainFunctionality: payload.RcPurposeMainFunctionality,
			RcResponsibility: payload.RcResponsibility,
			RcResponsibilityId: responsibilityIds ? (responsibilityIds.length === 1 ? responsibilityIds[0] : responsibilityIds) : undefined,
			RcResponsibilityEmail: resolvedResponsibility.email,
			RcResponsibilityLoginName: resolvedResponsibility.loginName,
			RcRemarks: payload.RcRemarks,
			DataType: payload.DataType
		};
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