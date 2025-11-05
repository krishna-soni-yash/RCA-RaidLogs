/* eslint-disable */
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SiteConfiguration } from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import IGenericService, { IFetchOptions, ISaveOptions, ISaveResult, IBatchSaveOptions, IUpdateOptions } from './IGenericServices';

export class GenericService implements IGenericService {
  private static DEFAULT_PAGE_SIZE = 2000;
  private static DEFAULT_MAX_RETRIES = 3;
  private static DEFAULT_RETRY_DELAY_MS = 800;

  private sp: any | null = null;
  private spInstances: Map<string, any> = new Map();

  constructor(spInstance?: any, context?: WebPartContext) {
    if (spInstance) {
      this.sp = spInstance;
    } else if (context) {
      this.sp = spfi().using(SPFx(context));
    }
  }

  public init(spInstance?: any, context?: WebPartContext) {
    if (spInstance) {
      this.sp = spInstance;
    } else if (context) {
      this.sp = spfi().using(SPFx(context));
    }
  }

  private isParentSiteList(listTitle: string): boolean {
    return SiteConfiguration.PARENT_LISTS.indexOf(listTitle) !== -1;
  }

  private getSiteUrlForList(listTitle: string, context: WebPartContext): string {
    if (this.isParentSiteList(listTitle)) {
      return context.pageContext.site?.absoluteUrl || context.pageContext.web.absoluteUrl;
    }

    return context.pageContext.web.absoluteUrl;
  }

  private getSpInstanceForSite(siteUrl: string, context: WebPartContext): any {
    if (this.spInstances.has(siteUrl)) {
      return this.spInstances.get(siteUrl);
    }

    let spInstance: any;
    if (siteUrl === context.pageContext.web.absoluteUrl) {
      spInstance = this.sp || spfi().using(SPFx(context));
    } else {
      spInstance = spfi(siteUrl).using(SPFx(context));
    }

    this.spInstances.set(siteUrl, spInstance);
    return spInstance;
  }

  private withRetry = async <T>(operation: () => Promise<T>, maxRetries = GenericService.DEFAULT_MAX_RETRIES, baseDelay = GenericService.DEFAULT_RETRY_DELAY_MS): Promise<T> => {
    let lastError: any;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await operation();
      } catch (err) {
        lastError = err;
        const wait = baseDelay * Math.pow(1.7, attempt - 1);
        await new Promise(res => setTimeout(res, wait));
      }
    }
    throw lastError;
  };

  

  public async fetchAllItems<T = any>(options: IFetchOptions): Promise<T[]> {
    if (!options) throw new Error(ErrorMessages.FETCH_OPTIONS_REQUIRED);

    const {
      context, spInstance, listTitle, select = [], expand = [], filter, orderBy,
      pageSize = GenericService.DEFAULT_PAGE_SIZE,
      maxRetries = GenericService.DEFAULT_MAX_RETRIES,
      retryDelayMs = GenericService.DEFAULT_RETRY_DELAY_MS,
      forceSiteUrl
    } = options;

    let targetSiteUrl: string;
    if (forceSiteUrl) {
      targetSiteUrl = forceSiteUrl;
    } else if (context) {
      targetSiteUrl = this.getSiteUrlForList(listTitle, context);
    } else {
      throw new Error(ErrorMessages.CONTEXT_REQUIRED_FOR_SITE);
    }

    let targetSp: any;
    if (spInstance) {
      targetSp = spInstance;
    } else if (context) {
      targetSp = this.getSpInstanceForSite(targetSiteUrl, context);
    } else {
      throw new Error(ErrorMessages.PNP_INSTANCE_NOT_INITIALIZED);
    }

    const effectivePageSize = Math.min(Math.max(1, pageSize), 5000);
    const list = targetSp.web.lists.getByTitle(listTitle);

    const buildItemsQuery = () => {
      let q: any = list.items;
      if (select && select.length > 0) q = q.select(...select);
      if (expand && expand.length > 0) q = q.expand(...expand);
      if (filter) q = q.filter(filter);
      if (orderBy) {
        const parts = orderBy.split(',').map(s => s.trim());
        parts.forEach(p => {
          const [field, dir] = p.split(' ').map(x => x.trim());
          q = q.orderBy(field, dir?.toLowerCase() === 'desc');
        });
      }
      q = q.top(effectivePageSize);
      return q;
    };

    const fetchItems = async (): Promise<T[]> => {
      return this.withRetry(async () => {
        const q = buildItemsQuery();
        return await q();
      }, maxRetries, retryDelayMs);
    };

    const results = await fetchItems();
    return results || [];
  }

  public async saveItem<T = any>(options: ISaveOptions): Promise<ISaveResult<T>> {
    if (!options) throw new Error(ErrorMessages.SAVE_OPTIONS_REQUIRED);

    const {
      context, spInstance, listTitle, item, itemId,
      select = [], expand = [],
      maxRetries = GenericService.DEFAULT_MAX_RETRIES,
      retryDelayMs = GenericService.DEFAULT_RETRY_DELAY_MS,
      forceSiteUrl,
      validateColumns = false
    } = options;

    if (!item || typeof item !== 'object') {
      throw new Error(ErrorMessages.INVALID_ITEM_DATA);
    }

    let targetSiteUrl: string;
    if (forceSiteUrl) {
      targetSiteUrl = forceSiteUrl;
    } else if (context) {
      targetSiteUrl = this.getSiteUrlForList(listTitle, context);
    } else {
      throw new Error(ErrorMessages.CONTEXT_REQUIRED_FOR_SITE);
    }

    let targetSp: any;
    if (spInstance) {
      targetSp = spInstance;
    } else if (context) {
      targetSp = this.getSpInstanceForSite(targetSiteUrl, context);
    } else {
      throw new Error(ErrorMessages.PNP_INSTANCE_NOT_INITIALIZED);
    }

    const list = targetSp.web.lists.getByTitle(listTitle);

    try {
      // Validate columns if requested
      if (validateColumns) {
        await this.validateItemColumns(targetSp, listTitle, item);
      }

      // Clean the item data (remove read-only fields, etc.)
      const cleanedItem = this.cleanItemForSave(item);

      const saveOperation = async (): Promise<ISaveResult<T>> => {
        return this.withRetry(async () => {
          let result: any;

          if (itemId) {
            // Update existing item
            const listItem = list.items.getById(itemId);
            result = await listItem.update(cleanedItem);

            // Fetch updated item if select/expand specified
            if (select.length > 0 || expand.length > 0) {
              let query = listItem;
              if (select.length > 0) query = query.select(...select);
              if (expand.length > 0) query = query.expand(...expand);
              const updatedItem = await query();

              return {
                success: true,
                item: updatedItem,
                itemId: itemId
              };
            }

            return {
              success: true,
              itemId: itemId
            };
          } else {
            // Create new item
            result = await list.items.add(cleanedItem);

            const resolveCreatedItemId = async (): Promise<number | undefined> => {
              if (result?.data) {
                const directId = result.data.Id ?? result.data.ID ?? result.data.id;
                if (directId !== undefined) {
                  const parsed = Number(directId);
                  return isNaN(parsed) ? undefined : parsed;
                }
              }

              if (result?.item) {
                try {
                  const itemData = await result.item.select('Id')();
                  const fallbackId = itemData?.Id ?? itemData?.ID ?? itemData?.id;
                  if (fallbackId !== undefined) {
                    const parsed = Number(fallbackId);
                    return isNaN(parsed) ? undefined : parsed;
                  }
                } catch (innerErr) {
                  console.warn('Unable to resolve created item ID from item.select', innerErr);
                }
              }

              return undefined;
            };

            const createdItemId = await resolveCreatedItemId();

            if (createdItemId === undefined) {
              console.warn('GenericService.saveItem: created item ID was not returned by SharePoint.');
              return {
                success: true,
                item: (result?.data as T | undefined)
              };
            }

            // Fetch created item if select/expand specified
            if (select.length > 0 || expand.length > 0) {
              let query = list.items.getById(createdItemId);
              if (select.length > 0) query = query.select(...select);
              if (expand.length > 0) query = query.expand(...expand);
              const createdItem = await query();

              return {
                success: true,
                item: createdItem,
                itemId: createdItemId
              };
            }

            return {
              success: true,
              item: result?.data,
              itemId: createdItemId
            };
          }
        }, maxRetries, retryDelayMs);
      };

      return await saveOperation();

    } catch (error: any) {
      console.error(`Error saving item to list ${listTitle}:`, error);
      return {
        success: false,
        error: error.message || 'Unknown error occurred'
      };
    }
  }

  public async saveBatchItems<T = any>(options: IBatchSaveOptions): Promise<ISaveResult<T>[]> {
    if (!options) throw new Error(ErrorMessages.BATCH_SAVE_OPTIONS_REQUIRED);

    const {
      context, spInstance, listTitle, items,
      batchSize = 100,
      maxRetries = GenericService.DEFAULT_MAX_RETRIES,
      retryDelayMs = GenericService.DEFAULT_RETRY_DELAY_MS,
      forceSiteUrl,
      validateColumns = false
    } = options;

    if (!items || items.length === 0) {
      return [];
    }

    let targetSiteUrl: string;
    if (forceSiteUrl) {
      targetSiteUrl = forceSiteUrl;
    } else if (context) {
      targetSiteUrl = this.getSiteUrlForList(listTitle, context);
    } else {
      throw new Error(ErrorMessages.CONTEXT_REQUIRED_FOR_SITE);
    }

    let targetSp: any;
    if (spInstance) {
      targetSp = spInstance;
    } else if (context) {
      targetSp = this.getSpInstanceForSite(targetSiteUrl, context);
    } else {
      throw new Error(ErrorMessages.PNP_INSTANCE_NOT_INITIALIZED);
    }

    const results: ISaveResult<T>[] = [];

    // Process items in batches
    for (let i = 0; i < items.length; i += batchSize) {
      const batch = items.slice(i, i + batchSize);

      const batchPromises = batch.map(item => this.saveItem<T>({
        context,
        spInstance: targetSp,
        listTitle,
        item,
        maxRetries,
        retryDelayMs,
        forceSiteUrl: targetSiteUrl,
        validateColumns
      }).catch(error => ({
        success: false,
        error: error.message || 'Batch item save failed'
      } as ISaveResult<T>)));

      const batchResults = await Promise.all(batchPromises);
      results.push(...batchResults);
    }

    return results;
  }

  public async updateItem<T = any>(options: IUpdateOptions): Promise<ISaveResult<T>> {
    if (!options) throw new Error(ErrorMessages.UPDATE_OPTIONS_REQUIRED);

    const {
      context, spInstance, listTitle, itemId, item,
      select = [], expand = [],
      maxRetries = GenericService.DEFAULT_MAX_RETRIES,
      retryDelayMs = GenericService.DEFAULT_RETRY_DELAY_MS,
      forceSiteUrl,
      validateColumns = false
    } = options;

    if (!itemId || itemId <= 0) {
      throw new Error(ErrorMessages.INVALID_ITEM_ID);
    }

    if (!item || typeof item !== 'object') {
      throw new Error(ErrorMessages.INVALID_ITEM_DATA);
    }

    let targetSiteUrl: string;
    if (forceSiteUrl) {
      targetSiteUrl = forceSiteUrl;
    } else if (context) {
      targetSiteUrl = this.getSiteUrlForList(listTitle, context);
    } else {
      throw new Error(ErrorMessages.CONTEXT_REQUIRED_FOR_SITE);
    }

    let targetSp: any;
    if (spInstance) {
      targetSp = spInstance;
    } else if (context) {
      targetSp = this.getSpInstanceForSite(targetSiteUrl, context);
    } else {
      throw new Error(ErrorMessages.PNP_INSTANCE_NOT_INITIALIZED);
    }

    const list = targetSp.web.lists.getByTitle(listTitle);

    try {
      // Validate columns if requested
      if (validateColumns) {
        await this.validateItemColumns(targetSp, listTitle, item);
      }

      // Clean the item data (remove read-only fields, etc.)
      const cleanedItem = this.cleanItemForSave(item);

      const updateOperation = async (): Promise<ISaveResult<T>> => {
        return this.withRetry(async () => {
          const listItem = list.items.getById(itemId);
          await listItem.update(cleanedItem);

          // Fetch updated item if select/expand specified
          if (select.length > 0 || expand.length > 0) {
            let query = listItem;
            if (select.length > 0) query = query.select(...select);
            if (expand.length > 0) query = query.expand(...expand);
            const updatedItem = await query();

            return {
              success: true,
              item: updatedItem,
              itemId: itemId
            };
          }

          return {
            success: true,
            itemId: itemId
          };
        }, maxRetries, retryDelayMs);
      };

      return await updateOperation();

    } catch (error: any) {
      console.error(`Error updating item ${itemId} in list ${listTitle}:`, error);
      return {
        success: false,
        error: error.message || 'Unknown error occurred'
      };
    }
  }

  private async validateItemColumns(sp: any, listTitle: string, item: any): Promise<void> {
    try {
      const list = sp.web.lists.getByTitle(listTitle);
      const fields = await list.fields.filter("Hidden eq false and ReadOnlyField eq false")();

      const fieldNames = fields.map((field: any) => field.InternalName);
      const itemKeys = Object.keys(item);

      const invalidFields = itemKeys.filter(key => !fieldNames.includes(key));

      if (invalidFields.length > 0) {
        console.warn(`Invalid fields for list ${listTitle}:`, invalidFields);
      }
    } catch (error) {
      console.warn(`Could not validate columns for list ${listTitle}:`, error);
    }
  }

  private cleanItemForSave(item: any): any {
    // Remove read-only and system fields that shouldn't be saved
    const fieldsToRemove = [
      'Id', 'ID', 'id',
      'Created', 'Modified',
      'AuthorId', 'EditorId',
      'Author', 'Editor',
      'GUID', 'UniqueId',
      'Version', '_ObjectVersion_',
      'FileSystemObjectType',
      'ServerRedirectedEmbedUri',
      'ServerRedirectedEmbedUrl',
      'ContentTypeId',
      '_ObjectIdentity_',
      '_ObjectType_',
      'odata.type',
      'odata.id',
      'odata.etag',
      'odata.editLink'
    ];

    const cleaned = { ...item };

    fieldsToRemove.forEach(field => {
      delete cleaned[field];
    });

    // Remove any fields starting with underscore (typically system fields)
    Object.keys(cleaned).forEach(key => {
      if (key.charAt(0) === '_' && key.charAt(key.length - 1) === '_') {
        delete cleaned[key];
      }
    });

    // Handle lookup fields - ensure they have proper format
    Object.keys(cleaned).forEach(key => {
      const value = cleaned[key];
      if (value && typeof value === 'object') {
        // Handle lookup field objects
        if (value.hasOwnProperty('Id') && key.indexOf('Id') !== key.length - 2) {
          cleaned[`${key}Id`] = value.Id;
          delete cleaned[key];
        }
        // Handle multi-choice or complex objects
        else if (Array.isArray(value)) {
          cleaned[key] = value;
        }
      }
    });

    return cleaned;
  }

  public clearCache(): void {
    this.spInstances.clear();
  }
}

export default new GenericService();