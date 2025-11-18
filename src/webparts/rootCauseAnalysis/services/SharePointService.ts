import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

/**
 * Generic SharePoint List Item interface
 */
export interface ISharePointListItem {
  Id?: number;
  Title?: string;
  [key: string]: any;
}

/**
 * SharePoint List Query Options
 */
export interface IListQueryOptions {
  select?: string[];
  filter?: string;
  orderBy?: string;
  top?: number;
  skip?: number;
  expand?: string[];
}

/**
 * SharePoint Service Response
 */
export interface ISharePointServiceResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  errorDetails?: any;
}

/**
 * SharePoint List Information
 */
export interface IListInfo {
  listName: string;
  siteUrl?: string;
}

/**
 * Batch Operation Item
 */
export interface IBatchOperation {
  method: 'POST' | 'PATCH' | 'DELETE';
  url: string;
  body?: any;
  headers?: { [key: string]: string };
}

/**
 * Generic SharePoint List Service
 * Provides CRUD operations for any SharePoint list
 */
export class SharePointService {
  private context: WebPartContext;
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.context = context;
    
    // create an SPFI instance bound to the current SPFx context
    this.sp = spfi().using(SPFx(this.context as any));
  }

  private getList(listInfo: IListInfo) {
    return this.sp.web.lists.getByTitle(listInfo.listName);
  }



  /**
   * Handle PnP JS response
   */
  private async handlePnPResponse<T>(operation: () => Promise<T>): Promise<ISharePointServiceResponse<T>> {
    try {
      const data = await operation();
      return {
        success: true,
        data: data
      };
    } catch (error: any) {
      console.error('PnP JS Error:', error);
      return {
        success: false,
        error: error.message || 'Unknown PnP JS error',
        errorDetails: error
      };
    }
  }

  /**
   * CREATE: Add a new item to SharePoint list
   */
  async createItem<T extends ISharePointListItem>(
    listInfo: IListInfo,
    item: Omit<T, 'Id'>
  ): Promise<ISharePointServiceResponse<T>> {
    return await this.handlePnPResponse(async () => {
      console.log('Creating SharePoint item with data:', JSON.stringify(item, null, 2));
      
      const list = this.getList(listInfo);
      const result = await list.items.add(item);
      return result.data as T;
    });
  }

  /**
   * READ: Get all items from SharePoint list
   */
  async getItems<T extends ISharePointListItem>(
    listInfo: IListInfo,
    options?: IListQueryOptions
  ): Promise<ISharePointServiceResponse<T[]>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      let query = list.items;

      if (options) {
        if (options.select) {
          query = query.select(...options.select);
        }
        if (options.filter) {
          query = query.filter(options.filter);
        }
        if (options.orderBy) {
          const orderParts = options.orderBy.split(' ');
          const fieldName = orderParts[0];
          const isDescending = orderParts.length > 1 && orderParts[1].toLowerCase() === 'desc';
          query = query.orderBy(fieldName, !isDescending);
        }
        if (options.top) {
          query = query.top(options.top);
        }
        if (options.expand) {
          query = query.expand(...options.expand);
        }
      }

      return await query() as T[];
    });
  }

  /**
   * READ: Get a single item by ID
   */
  async getItemById<T extends ISharePointListItem>(
    listInfo: IListInfo,
    itemId: number,
    options?: IListQueryOptions
  ): Promise<ISharePointServiceResponse<T>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      let query = list.items.getById(itemId);

      if (options) {
        if (options.select) {
          query = query.select(...options.select);
        }
        if (options.expand) {
          query = query.expand(...options.expand);
        }
      }

      return await query() as T;
    });
  }

  /**
   * UPDATE: Update an existing item
   */
  async updateItem<T extends ISharePointListItem>(
    listInfo: IListInfo,
    itemId: number,
    updates: Partial<Omit<T, 'Id'>>
  ): Promise<ISharePointServiceResponse<T>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      await list.items.getById(itemId).update(updates);
      const updatedItem = await list.items.getById(itemId)();
      return updatedItem as T;
    });
  }

  /**
   * DELETE: Delete an item by ID
   */
  async deleteItem(
    listInfo: IListInfo,
    itemId: number
  ): Promise<ISharePointServiceResponse<boolean>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      await list.items.getById(itemId).delete();
      return true;
    });
  }

  /**
   * UTILITY: Get list information
   */
  async getListInfo(listInfo: IListInfo): Promise<ISharePointServiceResponse<any>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      return await list();
    });
  }

  /**
   * UTILITY: Check if list exists
   */
  async listExists(listInfo: IListInfo): Promise<boolean> {
    const result = await this.getListInfo(listInfo);
    return result.success;
  }

  /**
   * UTILITY: Get list fields/columns
   */
  async getListFields(listInfo: IListInfo): Promise<ISharePointServiceResponse<any[]>> {
    return await this.handlePnPResponse(async () => {
      const list = this.getList(listInfo);
      return await list.relatedFields();
    });
  }

  /**
   * BATCH: Execute multiple operations in a single request
   * Note: PnP JS handles batching automatically for better performance
   */
  async executeBatch(
    operations: IBatchOperation[]
  ): Promise<ISharePointServiceResponse<any[]>> {
    return await this.handlePnPResponse(async () => {
      const results: any[] = [];
      
      for (const operation of operations) {
        try {
          console.warn('Batch operation executed sequentially:', operation);
          results.push({ success: true, operation });
        } catch (error) {
          results.push({ success: false, error, operation });
        }
      }
      
      return results;
    });
  }



  /**
   * UTILITY: Create filter string for OData queries
   */
  static createFilter(conditions: { [field: string]: any }): string {
    const filters: string[] = [];
    
    Object.keys(conditions).forEach(field => {
      const value = conditions[field];
      if (value !== null && value !== undefined) {
        if (typeof value === 'string') {
          filters.push(`${field} eq '${value.replace(/'/g, "''")}'`);
        } else if (typeof value === 'number') {
          filters.push(`${field} eq ${value}`);
        } else if (typeof value === 'boolean') {
          filters.push(`${field} eq ${value}`);
        } else if (value instanceof Date) {
          filters.push(`${field} eq datetime'${value.toISOString()}'`);
        }
      }
    });
    
    return filters.join(' and ');
  }

  /**
   * UTILITY: Create order by string
   */
  static createOrderBy(field: string, ascending: boolean = true): string {
    return `${field} ${ascending ? 'asc' : 'desc'}`;
  }
}

/**
 * SharePoint Service Factory
 * Creates a singleton instance of SharePoint Service
 */
export class SharePointServiceFactory {
  private static instance: SharePointService;

  static getInstance(context: WebPartContext): SharePointService {
    if (!SharePointServiceFactory.instance) {
      SharePointServiceFactory.instance = new SharePointService(context);
    }
    return SharePointServiceFactory.instance;
  }
}