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

	private extractNumberTokens(value: any): number[] {
		const results: number[] = [];
		const seen = new Set<number>();
		const consume = (input: any): void => {
			if (input === undefined || input === null) {
				return;
			}
			if (Array.isArray(input)) {
				input.forEach(consume);
				return;
			}
			if (typeof input === 'object') {
				if (Array.isArray((input as any).results)) {
					(input as any).results.forEach(consume);
					return;
				}
				consume((input as any).Id);
				consume((input as any).ID);
				consume((input as any).id);
				return;
			}
			if (typeof input === 'number' && !isNaN(input)) {
				if (!seen.has(input)) {
					seen.add(input);
					results.push(input);
				}
				return;
			}
			const text = String(input);
			if (!text) {
				return;
			}
			const sanitized = text.replace(/;#/g, ';');
			sanitized
				.split(/[;\s,]+/)
				.map(part => part.trim())
				.filter(part => part.length > 0)
				.forEach(part => {
					const parsed = Number(part);
					if (!isNaN(parsed) && !seen.has(parsed)) {
						seen.add(parsed);
						results.push(parsed);
					}
				});
		};
		consume(value);
		return results;
	}

	private extractStringTokens(value: any): string[] {
		const results: string[] = [];
		const seen = new Set<string>();
		const consume = (input: any): void => {
			if (input === undefined || input === null) {
				return;
			}
			if (Array.isArray(input)) {
				input.forEach(consume);
				return;
			}
			if (typeof input === 'object') {
				if (Array.isArray((input as any).results)) {
					(input as any).results.forEach(consume);
					return;
				}
				consume((input as any).Title);
				consume((input as any).Name);
				consume((input as any).DisplayName);
				consume((input as any).LoginName);
				consume((input as any).UserPrincipalName);
				consume((input as any).Email);
				consume((input as any).EMail);
				consume((input as any).SecondaryText);
				return;
			}
			const text = String(input);
			if (!text) {
				return;
			}
			const sanitized = text.replace(/;#/g, ';');
			sanitized
				.split(/[;\n]+/)
				.map(part => part.trim())
				.filter(part => part.length > 0)
				.forEach(part => {
					const key = part.toLowerCase();
					if (!seen.has(key)) {
						seen.add(key);
						results.push(part);
					}
				});
		};
		consume(value);
		return results;
	}

	private collapseNumberValues(values: Array<number | undefined>): number | number[] | undefined {
		const distinct: number[] = [];
		const seen = new Set<number>();
		for (const value of values) {
			if (typeof value === 'number' && !isNaN(value) && !seen.has(value)) {
				seen.add(value);
				distinct.push(value);
			}
		}
		if (distinct.length === 0) {
			return undefined;
		}
		return distinct.length === 1 ? distinct[0] : distinct;
	}

	private collapseStringValues(values: Array<string | undefined>): string | string[] | undefined {
		const distinct: string[] = [];
		const seen = new Set<string>();
		for (const value of values) {
			if (typeof value === 'string') {
				const trimmed = value.trim();
				if (trimmed.length > 0) {
					const key = trimmed.toLowerCase();
					if (!seen.has(key)) {
						seen.add(key);
						distinct.push(trimmed);
					}
				}
			}
		}
		if (distinct.length === 0) {
			return undefined;
		}
		return distinct.length === 1 ? distinct[0] : distinct;
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
		resolvedIds?: number | number[] | string | string[] | null,
		originalValue?: number | number[] | string | string[] | null,
		forceClear: boolean = false
	): number[] | undefined {
		const ids: number[] = [];
		ids.push(...this.normalizeIdArray(resolvedIds));
		ids.push(...this.normalizeIdArray(originalValue));
		const distinctIds: number[] = [];
		for (const id of ids) {
			if (typeof id === 'number' && !isNaN(id) && distinctIds.indexOf(id) === -1) {
				distinctIds.push(id);
			}
		}
		if (distinctIds.length > 0) {
			target[fieldName] = distinctIds;
			return distinctIds;
		}
		if (forceClear) {
			target[fieldName] = [];
			return [];
		}
		return undefined;
	}

	private normalizePersonField(item: any, fieldName: string): {
		displayName?: string;
		email?: string;
		loginName?: string;
		id?: number;
		displayNames: string[];
		emails: string[];
		loginNames: string[];
		ids: number[];
	} {
		const displayNames: string[] = [];
		const emails: string[] = [];
		const loginNames: string[] = [];
		const ids: number[] = [];
		const displayNameSet = new Set<string>();
		const emailSet = new Set<string>();
		const loginSet = new Set<string>();
		const idSet = new Set<number>();

		const pushNumbers = (value: any): void => {
			this.extractNumberTokens(value).forEach(num => {
				if (!idSet.has(num)) {
					idSet.add(num);
					ids.push(num);
				}
			});
		};

		const pushStrings = (collection: string[], seen: Set<string>, value: any, predicate?: (token: string) => boolean): void => {
			this.extractStringTokens(value).forEach(token => {
				if (!token) {
					return;
				}
				if (predicate && !predicate(token)) {
					return;
				}
				const key = token.toLowerCase();
				if (!seen.has(key)) {
					seen.add(key);
					collection.push(token);
				}
			});
		};

		const addCandidateObject = (candidate: any): void => {
			if (!candidate) {
				return;
			}
			pushNumbers(candidate?.Id);
			pushNumbers(candidate?.ID);
			pushNumbers(candidate?.id);
			pushStrings(displayNames, displayNameSet, [
				candidate?.Title,
				candidate?.Name,
				candidate?.DisplayName,
				candidate?.LoginName,
				candidate?.UserPrincipalName
			]);
			pushStrings(emails, emailSet, [
				candidate?.Email,
				candidate?.EMail,
				candidate?.SecondaryText
			], token => token.indexOf('@') !== -1);
			pushStrings(loginNames, loginSet, [
				candidate?.LoginName,
				candidate?.UserPrincipalName,
				candidate?.Name,
				candidate?.Email,
				candidate?.EMail
			]);
		};

		const addCandidate = (value: any): void => {
			if (value === undefined || value === null) {
				return;
			}
			if (Array.isArray(value)) {
				value.forEach(addCandidate);
				return;
			}
			if (typeof value === 'object') {
				if (Array.isArray((value as any).results)) {
					(value as any).results.forEach(addCandidate);
					return;
				}
				addCandidateObject(value);
				return;
			}
			if (typeof value === 'number') {
				pushNumbers(value);
				return;
			}
			if (typeof value === 'string') {
				pushNumbers(value);
				pushStrings(displayNames, displayNameSet, value);
				pushStrings(emails, emailSet, value, token => token.indexOf('@') !== -1);
				pushStrings(loginNames, loginSet, value);
			}
		};

		addCandidate(item?.[fieldName]);
		addCandidate(item?.[`${fieldName}Id`]);
		addCandidate(item?.[`${fieldName}Email`]);
		addCandidate(item?.[`${fieldName}LoginName`]);

		if (displayNames.length === 0 && typeof item?.[fieldName] === 'string') {
			pushStrings(displayNames, displayNameSet, item[fieldName]);
		}

		const primaryDisplay = displayNames.length > 0
			? displayNames[0]
			: this.pickFirstString([
				item?.[fieldName]
			]);
		const primaryEmail = emails.length > 0 ? emails[0] : this.pickFirstString([item?.[`${fieldName}Email`]]);
		const primaryLogin = loginNames.length > 0 ? loginNames[0] : this.pickFirstString([item?.[`${fieldName}LoginName`]]);
		const primaryId = ids.length > 0 ? ids[0] : this.pickFirstNumber([item?.[`${fieldName}Id`]]);

		return {
			displayName: primaryDisplay,
			email: primaryEmail,
			loginName: primaryLogin,
			id: typeof primaryId === 'number' && !isNaN(primaryId) ? primaryId : undefined,
			displayNames,
			emails,
			loginNames,
			ids
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
					const responsibilityDisplay = responsibilityInfo.displayNames.length > 0
						? responsibilityInfo.displayNames.join('; ')
						: (typeof it?.BpResponsibility === 'string' ? it.BpResponsibility : (responsibilityInfo.displayName ?? ''));
					const responsibilityIdValue = this.collapseNumberValues(responsibilityInfo.ids);
					const responsibilityEmailValue = this.collapseStringValues(responsibilityInfo.emails);
					const responsibilityLoginValue = this.collapseStringValues(responsibilityInfo.loginNames);

					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						BpBestPracticesDescription: it?.BpBestPracticesDescription ?? it?.BestPracticesDescription ?? it?.Description ?? it?.Title ?? '',
						BpReferences: it?.BpReferences ?? it?.References ?? '',
						BpResponsibility: responsibilityDisplay,
						BpResponsibilityId: responsibilityIdValue,
						BpResponsibilityEmail: responsibilityEmailValue,
						BpResponsibilityLoginName: responsibilityLoginValue,
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
		const responsibilityIds = this.assignMultiLookupId(payload, 'BpResponsibilityId', item.BpResponsibilityId, resolvedResponsibility.id);
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

		const responsibilityIdCandidates: number[] = [];
		if (Array.isArray(responsibilityIds)) {
			responsibilityIdCandidates.push(...responsibilityIds);
		}
		responsibilityIdCandidates.push(...this.extractNumberTokens(item.BpResponsibilityId));
		const responsibilityIdValue = this.collapseNumberValues(responsibilityIdCandidates);

		const emailCandidates = this.extractStringTokens(item.BpResponsibilityEmail);
		if (resolvedResponsibility.email) {
			emailCandidates.push(resolvedResponsibility.email);
		}
		const responsibilityEmailValue = this.collapseStringValues(emailCandidates);

		const loginCandidates = this.extractStringTokens(item.BpResponsibilityLoginName ?? item.BpResponsibilityEmail);
		if (resolvedResponsibility.loginName) {
			loginCandidates.push(resolvedResponsibility.loginName);
		}
		const responsibilityLoginValue = this.collapseStringValues(loginCandidates);

		const responsibilityDisplayTokens = this.extractStringTokens([
			item.BpResponsibility,
			resolvedResponsibility.displayName
		]);
		const responsibilityDisplay = responsibilityDisplayTokens.length > 0
			? responsibilityDisplayTokens.join('; ')
			: (resolvedResponsibility.displayName ?? (typeof item.BpResponsibility === 'string' ? item.BpResponsibility : ''));

		const savedItem: IBestPractices = {
			ID: hasValidId ? savedId : undefined,
			BpBestPracticesDescription: payload.BpBestPracticesDescription,
			BpReferences: payload.BpReferences,
			BpResponsibility: responsibilityDisplay,
			BpResponsibilityId: responsibilityIdValue,
			BpResponsibilityEmail: responsibilityEmailValue,
			BpResponsibilityLoginName: responsibilityLoginValue,
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
		const responsibilityIds = this.assignMultiLookupId(payload, 'BpResponsibilityId', item.BpResponsibilityId, resolvedResponsibility.id, true);
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

		const responsibilityIdCandidates: number[] = [];
		if (Array.isArray(responsibilityIds)) {
			responsibilityIdCandidates.push(...responsibilityIds);
		}
		responsibilityIdCandidates.push(...this.extractNumberTokens(item.BpResponsibilityId));
		const responsibilityIdValue = this.collapseNumberValues(responsibilityIdCandidates);

		const emailCandidates = this.extractStringTokens(item.BpResponsibilityEmail);
		if (resolvedResponsibility.email) {
			emailCandidates.push(resolvedResponsibility.email);
		}
		const responsibilityEmailValue = this.collapseStringValues(emailCandidates);

		const loginCandidates = this.extractStringTokens(item.BpResponsibilityLoginName ?? item.BpResponsibilityEmail);
		if (resolvedResponsibility.loginName) {
			loginCandidates.push(resolvedResponsibility.loginName);
		}
		const responsibilityLoginValue = this.collapseStringValues(loginCandidates);

		const responsibilityDisplayTokens = this.extractStringTokens([
			item.BpResponsibility,
			resolvedResponsibility.displayName
		]);
		const responsibilityDisplay = responsibilityDisplayTokens.length > 0
			? responsibilityDisplayTokens.join('; ')
			: (resolvedResponsibility.displayName ?? (typeof item.BpResponsibility === 'string' ? item.BpResponsibility : ''));

		return {
			ID: item.ID,
			BpBestPracticesDescription: payload.BpBestPracticesDescription,
			BpReferences: payload.BpReferences,
			BpResponsibility: responsibilityDisplay,
			BpResponsibilityId: responsibilityIdValue,
			BpResponsibilityEmail: responsibilityEmailValue,
			BpResponsibilityLoginName: responsibilityLoginValue,
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
					const responsibilityDisplay = responsibilityInfo.displayNames.length > 0
						? responsibilityInfo.displayNames.join('; ')
						: (typeof it?.RcResponsibility === 'string' ? it.RcResponsibility : (responsibilityInfo.displayName ?? ''));
					const responsibilityIdValue = this.collapseNumberValues(responsibilityInfo.ids);
					const responsibilityEmailValue = this.collapseStringValues(responsibilityInfo.emails);
					const responsibilityLoginValue = this.collapseStringValues(responsibilityInfo.loginNames);

					return {
						ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
						RcComponentName: it?.RcComponentName ?? it?.ComponentName ?? it?.Title ?? '',
						RcLocation: it?.RcLocation ?? it?.Location ?? '',
						RcPurposeMainFunctionality: it?.RcPurposeMainFunctionality ?? it?.Purpose ?? '',
						RcResponsibility: responsibilityDisplay,
						RcResponsibilityId: responsibilityIdValue,
						RcResponsibilityEmail: responsibilityEmailValue,
						RcResponsibilityLoginName: responsibilityLoginValue,
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
		const responsibilityIds = this.assignMultiLookupId(payload, 'RcResponsibilityId', item.RcResponsibilityId, resolvedResponsibility.id);
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

		const responsibilityIdCandidates: number[] = [];
		if (Array.isArray(responsibilityIds)) {
			responsibilityIdCandidates.push(...responsibilityIds);
		}
		responsibilityIdCandidates.push(...this.extractNumberTokens(item.RcResponsibilityId));
		const responsibilityIdValue = this.collapseNumberValues(responsibilityIdCandidates);

		const emailCandidates = this.extractStringTokens(item.RcResponsibilityEmail);
		if (resolvedResponsibility.email) {
			emailCandidates.push(resolvedResponsibility.email);
		}
		const responsibilityEmailValue = this.collapseStringValues(emailCandidates);

		const loginCandidates = this.extractStringTokens(item.RcResponsibilityLoginName ?? item.RcResponsibilityEmail);
		if (resolvedResponsibility.loginName) {
			loginCandidates.push(resolvedResponsibility.loginName);
		}
		const responsibilityLoginValue = this.collapseStringValues(loginCandidates);

		const responsibilityDisplayTokens = this.extractStringTokens([
			item.RcResponsibility,
			resolvedResponsibility.displayName
		]);
		const responsibilityDisplay = responsibilityDisplayTokens.length > 0
			? responsibilityDisplayTokens.join('; ')
			: (resolvedResponsibility.displayName ?? (typeof item.RcResponsibility === 'string' ? item.RcResponsibility : ''));

		const savedItem: IReusableComponents = {
			ID: hasValidId ? savedId : undefined,
			RcComponentName: payload.RcComponentName,
			RcLocation: payload.RcLocation,
			RcPurposeMainFunctionality: payload.RcPurposeMainFunctionality,
			RcResponsibility: responsibilityDisplay,
			RcResponsibilityId: responsibilityIdValue,
			RcResponsibilityEmail: responsibilityEmailValue,
			RcResponsibilityLoginName: responsibilityLoginValue,
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
		const responsibilityIds = this.assignMultiLookupId(payload, 'RcResponsibilityId', item.RcResponsibilityId, resolvedResponsibility.id, true);
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

		const responsibilityIdCandidates: number[] = [];
		if (Array.isArray(responsibilityIds)) {
			responsibilityIdCandidates.push(...responsibilityIds);
		}
		responsibilityIdCandidates.push(...this.extractNumberTokens(item.RcResponsibilityId));
		const responsibilityIdValue = this.collapseNumberValues(responsibilityIdCandidates);

		const emailCandidates = this.extractStringTokens(item.RcResponsibilityEmail);
		if (resolvedResponsibility.email) {
			emailCandidates.push(resolvedResponsibility.email);
		}
		const responsibilityEmailValue = this.collapseStringValues(emailCandidates);

		const loginCandidates = this.extractStringTokens(item.RcResponsibilityLoginName ?? item.RcResponsibilityEmail);
		if (resolvedResponsibility.loginName) {
			loginCandidates.push(resolvedResponsibility.loginName);
		}
		const responsibilityLoginValue = this.collapseStringValues(loginCandidates);

		const responsibilityDisplayTokens = this.extractStringTokens([
			item.RcResponsibility,
			resolvedResponsibility.displayName
		]);
		const responsibilityDisplay = responsibilityDisplayTokens.length > 0
			? responsibilityDisplayTokens.join('; ')
			: (resolvedResponsibility.displayName ?? (typeof item.RcResponsibility === 'string' ? item.RcResponsibility : ''));

		return {
			ID: item.ID,
			RcComponentName: payload.RcComponentName,
			RcLocation: payload.RcLocation,
			RcPurposeMainFunctionality: payload.RcPurposeMainFunctionality,
			RcResponsibility: responsibilityDisplay,
			RcResponsibilityId: responsibilityIdValue,
			RcResponsibilityEmail: responsibilityEmailValue,
			RcResponsibilityLoginName: responsibilityLoginValue,
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