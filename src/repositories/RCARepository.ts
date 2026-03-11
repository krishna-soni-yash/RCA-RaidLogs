import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRCAList } from '../models/IRCAList';
import ErrorMessages from '../common/ErrorMessages';
import { SubSiteListNames, selectedFields, expandFields } from '../common/Constants';
import IRCARepository from '../repositories/repositoriesInterface/IRCARepository';
import RCAEmailTriggerService from '../services/RCAEmailTriggerService';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
//import getSpInstanceForSite from '../services/GenericServices';
//import getSiteUrlForList from '../services/GenericServices';

//import { SPHttpClient } from '@microsoft/sp-http'; // << added import


/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class RCARepository implements IRCARepository {
    private service: IGenericService;
    private cache: IRCAList[] | null = null;
    private cacheTimestamp = 0;
    private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

    private serializePeopleField(raw: any, rawId?: any): string {
        if (!raw && rawId === undefined) {
            return '';
        }

        const values: string[] = [];
        const seenIds: string[] = [];
        const pendingContacts: string[] = [];

        const addValue = (idValue: any, contactValue: any) => {
            const idStr = String(idValue ?? '').trim();
            const contactStr = String(contactValue ?? '').trim();
            if (!idStr && !contactStr) {
                return;
            }

            if (idStr) {
                if (seenIds.indexOf(idStr) !== -1) {
                    if (contactStr.length > 0) {
                        for (let i = 0; i < values.length; i++) {
                            if (values[i].indexOf(`${idStr}|`) === 0) {
                                const existing = values[i].split('|');
                                const currentContact = existing.length > 1 ? existing[1] : '';
                                if (!currentContact.length) {
                                    values[i] = `${idStr}|${contactStr}`;
                                }
                                break;
                            }
                        }
                    }
                    return;
                }
                values.push(`${idStr}|${contactStr}`.trim());
                seenIds.push(idStr);
            } else if (contactStr.length > 0) {
                values.push(`|${contactStr}`);
            }
        };

        const extractFromUser = (user: any) => {
            if (!user || typeof user !== 'object') {
                return;
            }
            const idCandidate = user.Id ?? user.ID ?? user.id ?? user.Key ?? user.key;
            const contactCandidate = user.EMail ?? user.Email ?? user.email ?? user.UserPrincipalName ?? user.userPrincipalName ?? user.LoginName ?? user.loginName ?? user.Title ?? user.text ?? user.DisplayName ?? user.Name ?? '';
            if (idCandidate !== undefined && idCandidate !== null && String(idCandidate).trim().length > 0) {
                addValue(idCandidate, contactCandidate);
            } else if (contactCandidate && String(contactCandidate).trim().length > 0) {
                pendingContacts.push(String(contactCandidate).trim());
            }
        };

        if (Array.isArray(raw)) {
            raw.forEach(extractFromUser);
        } else if (raw && typeof raw === 'object') {
            if (Array.isArray(raw.results)) {
                raw.results.forEach(extractFromUser);
            } else {
                extractFromUser(raw);
            }
        } else if (typeof raw === 'string' && raw.trim().length > 0) {
            return raw.trim();
        }

        const idList: string[] = [];
        if (rawId !== undefined && rawId !== null) {
            if (Array.isArray(rawId)) {
                rawId.forEach((id: any) => {
                    const idStr = String(id ?? '').trim();
                    if (idStr.length > 0) {
                        idList.push(idStr);
                    }
                });
            } else if (typeof rawId === 'object' && Array.isArray(rawId.results)) {
                rawId.results.forEach((id: any) => {
                    const idStr = String(id ?? '').trim();
                    if (idStr.length > 0) {
                        idList.push(idStr);
                    }
                });
            } else {
                const singleId = String(rawId ?? '').trim();
                if (singleId.length > 0) {
                    idList.push(singleId);
                }
            }
        }

        if (idList.length > 0) {
            idList.forEach((id, index) => {
                const contact = index < pendingContacts.length ? pendingContacts[index] : '';
                addValue(id, contact);
            });
        }

        if (values.length === 0 && pendingContacts.length > 0) {
            pendingContacts.forEach(contact => addValue('', contact));
        }

        if (values.length === 0) {
            return '';
        }

        const uniqueValues: string[] = [];
        for (let i = 0; i < values.length; i++) {
            const trimmed = values[i].trim();
            if (!trimmed) continue;
            if (uniqueValues.indexOf(trimmed) === -1) {
                uniqueValues.push(trimmed);
            }
        }
        return uniqueValues.join('; ');
    }

    private parsePeopleIds(raw: any): number[] {
        const ids: number[] = [];
        if (raw === undefined || raw === null) return ids;

        const pushIfValid = (n: any) => {
            const num = Number(n);
            if (!isNaN(num) && num > 0) ids.push(num);
        };

        if (Array.isArray(raw)) {
            raw.forEach(r => {
                if (typeof r === 'number') pushIfValid(r);
                else if (typeof r === 'string') {
                    const left = (r || '').split('|')[0].trim();
                    pushIfValid(left || r);
                } else if (r && typeof r === 'object') {
                    const idCandidate = r.Id ?? r.ID ?? r.id ?? r.Key ?? r.key;
                    pushIfValid(idCandidate);
                }
            });
            const unique = ids.filter((v, i, a) => a.indexOf(v) === i);
            return unique;
        }

        if (typeof raw === 'number') { pushIfValid(raw); return ids; }
        if (typeof raw === 'string') {
            const parts = raw.split(/; ?/);
            parts.forEach(p => {
                const left = (p || '').split('|')[0].trim();
                pushIfValid(left);
            });
            const unique = ids.filter((v, i, a) => a.indexOf(v) === i);
            return unique;
        }

        if (raw && typeof raw === 'object') {
            const idCandidate = raw.Id ?? raw.ID ?? raw.id ?? raw.Key ?? raw.key;
            pushIfValid(idCandidate);
        }

        const unique = ids.filter((v, i, a) => a.indexOf(v) === i);
        return unique;
    }

    constructor(service?: IGenericService) {
        this.service = service ?? genericService;

    }

    public setService(service: IGenericService): void {
        this.service = service;
    }
    // private normalizeSiteUrl(value?: string): string {
    //     return (value || '').trim().replace(/\/+$/, '').toLowerCase();
    //   }
    public async getRCAItems(useCache: boolean = true, context?: WebPartContext): Promise<IRCAList[]> {
        const now = Date.now();

        if (useCache && this.cache && (now - this.cacheTimestamp) < this.CACHE_DURATION) {
            return this.cache;
        }

        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }

        try {
            const tableSelectFields: string[] = [
                'Id',
                'Title',
                'CauseCategory',
                'RCAPriority',
                'RootCause',
                'RCATypeOfAction',
                'ActionPlanCorrection',
                'ResponsibilityCorrection/Id',
                'ResponsibilityCorrection/Title',
                'ResponsibilityCorrection/EMail',
                'ResponsibilityCorrectionId',
                'PlannedClosureDateCorrection',
                'ActualClosureDateCorrection',
                'ActionPlanCorrective',
                'ResponsibilityCorrective/Id',
                'ResponsibilityCorrective/Title',
                'ResponsibilityCorrective/EMail',
                'ResponsibilityCorrectiveId',
                'PlannedClosureDateCorrective',
                'ActualClosureDateCorrective',
                'ActionPlanPreventive',
                'ResponsibilityPreventive/Id',
                'ResponsibilityPreventive/Title',
                'ResponsibilityPreventive/EMail',
                'ResponsibilityPreventiveId',
                'PlannedClosureDatePreventive',
                'ActualClosureDatePreventive',
                'Modified'
            ];

            const tableExpandFields: string[] = [
                'ResponsibilityCorrection',
                'ResponsibilityCorrective',
                'ResponsibilityPreventive'
            ];

            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);

            const items = await this.service.fetchAllItems<any>({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                select: tableSelectFields,
                pageSize: 2000,
                expand: tableExpandFields
                //filter: 'IsActive eq 1 and ProjectType/Title eq \'' + (selectedProjectType) + '\'',
                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

            });

            const normalized = (items || []).map((it: any) => ({
                ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
                LinkTitle: it?.Title || it?.LinkTitle || '',
                ProblemStatementNumber: it?.ProblemStatementNumber || '',
                ProblemStatement: it?.ProblemStatement || '',
                CauseCategory: it?.CauseCategory || '',
                RCASource: it?.RCASource || '',
                RCAPriority: it?.RCAPriority || '',
                RelatedMetric: it?.RelatedMetric || '',
                Cause: it?.Cause || '',
                RootCause: it?.RootCause || '',
                RCATechniqueUsedAndReference: it?.RCATechniqueUsedAndReference || '',
                RCATypeOfAction: it?.RCATypeOfAction || '',

                ActionPlanCorrection: it?.ActionPlanCorrection || '',
                ResponsibilityCorrection: this.serializePeopleField(it?.ResponsibilityCorrection, it?.ResponsibilityCorrectionId),
                PlannedClosureDateCorrection: it?.PlannedClosureDateCorrection || '',
                ActualClosureDateCorrection: it?.ActualClosureDateCorrection || '',

                ActionPlanCorrective: it?.ActionPlanCorrective || '',
                ResponsibilityCorrective: this.serializePeopleField(it?.ResponsibilityCorrective, it?.ResponsibilityCorrectiveId),
                PlannedClosureDateCorrective: it?.PlannedClosureDateCorrective || '',
                ActualClosureDateCorrective: it?.ActualClosureDateCorrective || '',

                ActionPlanPreventive: it?.ActionPlanPreventive || '',
                ResponsibilityPreventive: this.serializePeopleField(it?.ResponsibilityPreventive, it?.ResponsibilityPreventiveId),
                PlannedClosureDatePreventive: it?.PlannedClosureDatePreventive || '',
                ActualClosureDatePreventive: it?.ActualClosureDatePreventive || '',

                PerformanceBeforeActionPlan: it?.PerformanceBeforeActionPlan || '',
                PerformanceAfterActionPlan: it?.PerformanceAfterActionPlan || '',
                QuantitativeOrStatisticalEffecti: it?.Quantitative_x0020_Or_x0020_Stat || '',
                Remarks: it?.Remarks || '',
                RelatedSubMetric: it?.RelatedSubMetric || '',
                Modified: it?.Modified || '',
                attachments: []
            })) as unknown as IRCAList[];

            this.cache = normalized;
            this.cacheTimestamp = now;

            return this.cache;
        } catch (error: any) {
            throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
        }
    }

  

    public async saveRCAItem(item: IRCAList, context?: WebPartContext): Promise<any> {
        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }
        try {
            const payload: Record<string, unknown> = {};
            if (item.LinkTitle !== undefined) payload.Title = item.LinkTitle;
            if (item.ProblemStatementNumber !== undefined) payload.ProblemStatementNumber = item.ProblemStatementNumber;
            if (item.CauseCategory !== undefined) payload.CauseCategory = item.CauseCategory;
            if (item.RCASource !== undefined) payload.RCASource = item.RCASource;
            if (item.RCAPriority !== undefined) payload.RCAPriority = item.RCAPriority;
            if (item.RelatedMetric !== undefined) payload.RelatedMetric = item.RelatedMetric;
            if (item.Cause !== undefined) payload.Cause = item.Cause;
            if (item.RootCause !== undefined) payload.RootCause = item.RootCause;
            if (item.RCATechniqueUsedAndReference !== undefined) payload.RCATechniqueUsedAndReference = item.RCATechniqueUsedAndReference;
            if (item.RCATypeOfAction !== undefined) {
                try {
                    const raw = item.RCATypeOfAction;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(',') : []);

                    const values: string[] = parts
                        .map(p => p.trim())
                        .filter(p => p.length > 0);

                    if (values.length > 0) {
                        payload.RCATypeOfAction = values ;
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid values parsed for RCATypeOfAction; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse RCATypeOfAction', e);
                }
            }

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;
            if (item.ResponsibilityCorrection !== undefined) {
                try {
                    const raw = item.ResponsibilityCorrection;
                    console.log('RCARepository.saveRCAItem - ResponsibilityCorrection raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.saveRCAItem - ResponsibilityCorrection parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityCorrectionId = ids;
                        console.log('RCARepository.saveRCAItem - ResponsibilityCorrectionId payload:', payload.ResponsibilityCorrectionId);
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityCorrection', e);
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                try {
                    const raw = item.ResponsibilityCorrective;
                    console.log('RCARepository.saveRCAItem - ResponsibilityCorrective raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.saveRCAItem - ResponsibilityCorrective parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityCorrectiveId = ids;
                        console.log('RCARepository.saveRCAItem - ResponsibilityCorrectiveId payload:', payload.ResponsibilityCorrectiveId);
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityCorrective; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityCorrective', e);
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                try {
                    const raw = item.ResponsibilityPreventive;
                    console.log('RCARepository.saveRCAItem - ResponsibilityPreventive raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.saveRCAItem - ResponsibilityPreventive parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityPreventiveId = ids;
                        console.log('RCARepository.saveRCAItem - ResponsibilityPreventiveId payload:', payload.ResponsibilityPreventiveId);
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityPreventive; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityPreventive', e);
                }
            }
            if (item.PlannedClosureDatePreventive !== undefined) payload.PlannedClosureDatePreventive = item.PlannedClosureDatePreventive;
            if (item.ActualClosureDatePreventive !== undefined) payload.ActualClosureDatePreventive = item.ActualClosureDatePreventive;
            if (item.RelatedSubMetric !== undefined) payload.RelatedSubMetric = item.RelatedSubMetric;
            if (item.PerformanceBeforeActionPlan !== undefined) payload.PerformanceBeforeActionPlan = item.PerformanceBeforeActionPlan;
            if (item.PerformanceAfterActionPlan !== undefined) payload.PerformanceAfterActionPlan = item.PerformanceAfterActionPlan;
            if (item.QuantitativeOrStatisticalEffecti !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = item.QuantitativeOrStatisticalEffecti;
            if (item.Remarks !== undefined) payload.Remarks = item.Remarks;

            // map performance / quantitative / remarks fields for update
            if (item.PerformanceBeforeActionPlan !== undefined) payload.PerformanceBeforeActionPlan = item.PerformanceBeforeActionPlan;
            if (item.PerformanceAfterActionPlan !== undefined) payload.PerformanceAfterActionPlan = item.PerformanceAfterActionPlan;
            if ((item as any).Quantitative_x0020_Or_x0020_Stat !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = (item as any).Quantitative_x0020_Or_x0020_Stat;
            else if ((item as any).QuantitativeOrStatisticalEffecti !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = (item as any).QuantitativeOrStatisticalEffecti;
            if (item.Remarks !== undefined) payload.Remarks = item.Remarks;
            // map performance / quantitative / remarks fields
            if (item.PerformanceBeforeActionPlan !== undefined) payload.PerformanceBeforeActionPlan = item.PerformanceBeforeActionPlan;
            if (item.PerformanceAfterActionPlan !== undefined) payload.PerformanceAfterActionPlan = item.PerformanceAfterActionPlan;
            // support both possible property names coming from form/repo
            if ((item as any).Quantitative_x0020_Or_x0020_Stat !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = (item as any).Quantitative_x0020_Or_x0020_Stat;
            else if ((item as any).QuantitativeOrStatisticalEffecti !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = (item as any).QuantitativeOrStatisticalEffecti;
            if (item.Remarks !== undefined) payload.Remarks = item.Remarks;

            console.log('RCARepository.saveRCAItem - Final payload:', JSON.stringify(payload, null, 2));
            const result = await this.service.saveItem<IRCAList>({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                item: payload,
                select: selectedFields,
                expand: expandFields
            });
            this.refresh();

            // Attempt to create RCAEmailTrigger entry (best-effort, do not block save)
            try {
                const createdId = (result && result.itemId) ? result.itemId : (result && result.item && (result.item.ID));
                const emailService = new RCAEmailTriggerService(context);
                const createdItemForEmail: any = item;
                createdItemForEmail.ID = createdId;
                // fire-and-forget but await to surface errors in logs
                await emailService.createEmailTrigger(createdItemForEmail);
            } catch (e) {
                console.warn('Failed to create RCAEmailTrigger entry:', e);
            }

            return result;
        } catch (error: any) {
            throw new Error('Failed to save RCA item: ' + (error?.message || error));
        }
    }

    public async updateRCAItem(itemId: number, item: IRCAList, context?: WebPartContext): Promise<any> {
        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }
        try {
            const payload: Record<string, unknown> = {};
            if (item.LinkTitle !== undefined) payload.Title = item.LinkTitle;
            if (item.ProblemStatementNumber !== undefined) payload.ProblemStatementNumber = item.ProblemStatementNumber;
            if (item.CauseCategory !== undefined) payload.CauseCategory = item.CauseCategory;
            if (item.RCASource !== undefined) payload.RCASource = item.RCASource;
            if (item.RCAPriority !== undefined) payload.RCAPriority = item.RCAPriority;
            if (item.RelatedMetric !== undefined) payload.RelatedMetric = item.RelatedMetric;
            if (item.Cause !== undefined) payload.Cause = item.Cause;
            if (item.RootCause !== undefined) payload.RootCause = item.RootCause;
            if (item.RCATechniqueUsedAndReference !== undefined) payload.RCATechniqueUsedAndReference = item.RCATechniqueUsedAndReference;
            if (item.RCATypeOfAction !== undefined) {
                try {
                    const raw = item.RCATypeOfAction;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(',') : []);

                    const values: string[] = parts
                        .map(p => p.trim())
                        .filter(p => p.length > 0);

                    if (values.length > 0) {
                        payload.RCATypeOfAction = values;
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid values parsed for RCATypeOfAction; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse RCATypeOfAction', e);
                }
            }

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;

            if (item.ResponsibilityCorrection !== undefined) {
                try {
                    const raw = item.ResponsibilityCorrection;
                    console.log('RCARepository.updateRCAItem - ResponsibilityCorrection raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.updateRCAItem - ResponsibilityCorrection parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityCorrectionId = ids;
                        console.log('RCARepository.updateRCAItem - ResponsibilityCorrectionId payload:', payload.ResponsibilityCorrectionId);
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityCorrection', e);
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                try {
                    const raw = item.ResponsibilityCorrective;
                    console.log('RCARepository.updateRCAItem - ResponsibilityCorrective raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.updateRCAItem - ResponsibilityCorrective parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityCorrectiveId = ids;
                        console.log('RCARepository.updateRCAItem - ResponsibilityCorrectiveId payload:', payload.ResponsibilityCorrectiveId);
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityCorrective; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityCorrective', e);
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                try {
                    const raw = item.ResponsibilityPreventive;
                    console.log('RCARepository.updateRCAItem - ResponsibilityPreventive raw:', raw);
                    const ids = this.parsePeopleIds(raw);
                    console.log('RCARepository.updateRCAItem - ResponsibilityPreventive parsed IDs:', ids);
                    if (ids.length > 0) {
                        payload.ResponsibilityPreventiveId = ids;
                        console.log('RCARepository.updateRCAItem - ResponsibilityPreventiveId payload:', payload.ResponsibilityPreventiveId);
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityPreventive; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityPreventive', e);
                }
            }
            
            if (item.PlannedClosureDatePreventive !== undefined) payload.PlannedClosureDatePreventive = item.PlannedClosureDatePreventive;
            if (item.ActualClosureDatePreventive !== undefined) payload.ActualClosureDatePreventive = item.ActualClosureDatePreventive;
            if (item.RelatedSubMetric !== undefined) payload.RelatedSubMetric = item.RelatedSubMetric;

            
            if (item.PerformanceBeforeActionPlan !== undefined) payload.PerformanceBeforeActionPlan = item.PerformanceBeforeActionPlan;
            if (item.PerformanceAfterActionPlan !== undefined) payload.PerformanceAfterActionPlan = item.PerformanceAfterActionPlan;
            if (item.QuantitativeOrStatisticalEffecti !== undefined) payload.Quantitative_x0020_Or_x0020_Stat = item.QuantitativeOrStatisticalEffecti;
            if (item.Remarks !== undefined) payload.Remarks = item.Remarks;

            console.log('RCARepository.updateRCAItem - Final payload:', JSON.stringify(payload, null, 2));
            await this.service.updateItem({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                itemId: itemId,
                item: payload,
                select: selectedFields,
                expand: expandFields
            });
            // clear cached items so subsequent reads reflect the updated values
            this.refresh();
        }
        catch (error: any) {
            throw new Error('Failed to update RCA item: ' + (error?.message || error));
        }
    }
    public async uploadRCAAttachment(itemId: number, file: File, context?: WebPartContext): Promise<void> {
        if (!context || !itemId || !file) return;
        const targetSiteUrl = await this.service.getSiteUrlForList(SubSiteListNames.RootCauseAnalysis, context);
        const sp = await this.service.getSpInstanceForSite(targetSiteUrl, context);
        const item = sp?.web?.lists?.getByTitle?.(SubSiteListNames.RootCauseAnalysis)?.items?.getById?.(itemId);
        if (!item || !item.attachmentFiles || typeof item.attachmentFiles.add !== 'function') {
            console.error('uploadRCAAttachment: attachmentFiles.add not available', { hasItem: !!item });
            return;
        }
        await item.attachmentFiles.add(file.name, file);
    }
    public async deleteRCAAttachment(itemId: number, fileName: string, context?: WebPartContext): Promise<void> {
        if (!context || !itemId || !fileName) return;
        const targetSiteUrl = await this.service.getSiteUrlForList(SubSiteListNames.RootCauseAnalysis, context);
        const sp = await this.service.getSpInstanceForSite(targetSiteUrl, context);
        const attachment = sp?.web?.lists?.getByTitle?.(SubSiteListNames.RootCauseAnalysis)?.items?.getById?.(itemId)?.attachmentFiles?.getByName?.(fileName);
        if (!attachment || typeof attachment.delete !== 'function') {
            console.error('deleteRCAAttachment: attachmentFiles.getByName.delete not available', { hasAttachment: !!attachment });
            return;
        }
        await attachment.delete();
    }

    public refresh(): void {
        this.cache = null;
        this.cacheTimestamp = 0;
    }

    public getCacheStatus(): { cached: boolean; itemCount: number; age: number } {
        const now = Date.now();
        return {
            cached: this.cache !== null,
            itemCount: this.cache?.length || 0,
            age: this.cache ? now - this.cacheTimestamp : 0
        };
    }
}

const defaultInstance = new RCARepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getRCAItems = async (useCache: boolean = false, context?: WebPartContext): Promise<IRCAList[]> => defaultInstance.getRCAItems(useCache, context);
export const saveRCAItem = async (item: IRCAList, context?: WebPartContext): Promise<any> => defaultInstance.saveRCAItem(item, context);
export const updateRCAItem = async (itemId: number, item: IRCAList, context?: WebPartContext): Promise<any> => defaultInstance.updateRCAItem(itemId, item, context);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();
export const deleteRCAAttachment = async (itemId: number, fileName: string, context?: WebPartContext): Promise<void> => defaultInstance.deleteRCAAttachment(itemId, fileName, context);
export const uploadRCAAttachment = async (itemId: number, file: File, context?: WebPartContext): Promise<void> => defaultInstance.uploadRCAAttachment(itemId, file, context);



