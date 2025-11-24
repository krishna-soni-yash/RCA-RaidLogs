import { WebPartContext } from '@microsoft/sp-webpart-base';
import GenericServiceInstance from '../../../../services/GenericServices';
import { IGenericService } from '../../../../services/IGenericServices';
import { IRaidItem, RaidType, IRaidAction } from './interfaces/IRaidItem';
import { LIST_NAMES } from '../../../../common/Constants';
import { 
  IExtendedRaidItem
} from './interfaces/IRaidService';

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
 * SharePoint List Item interface for RAID items
 * Extends the generic interface with RAID-specific fields
 * Field names match SharePoint internal field names
 */
export interface IRaidSharePointItem extends ISharePointListItem {
  Id?: number;
  Title?: string;
  SelectType: RaidType; // Internal name for RAID Type selection
  RAIDId?: string; // Unique identifier for grouping Risk items with their actions
  IdentificationDate?: string;
  RiskDescription?: string; // Internal name for Description
  AssociatedGoal?: string;
  RiskSource?: string; // Internal name for Source
  RiskCategory?: string; // Internal name for Category
  Impact?: string;
  RiskPriority?: string; // Internal name for Priority
  
  // Risk specific fields
  ImpactValue?: number;
  ProbabilityValue?: number;
  RiskExposure?: number;
  
  // Opportunity specific fields
  PotentialCost?: number;
  PotentialBenefit?: number;
  OpportunityValue?: number;
  TypeOfAction?: string;
  ActionPlan?: string;
  Responsibility?: any; // Expanded people picker field (when using expand in query)
  ResponsibilityId?: any; // People picker field - can be string format "id|email; id|email" or numeric
  TargetDate?: string;
  ActualDate?: string;
  RiskStatus?: string; // Internal name for Status
  
  // Issue/Assumption/Dependency/Constraints specific fields
  IssueDetails?: string; // Internal name for Details
  IDADate?: string; // Internal name for Date
  ByWhom?: any; // Expanded people picker field (when using expand in query)
  ByWhomId?: any; // People picker field - can be string format "id|email; id|email" or numeric
  ImplementationActions?: string;
  PlannedClosureDate?: string;
  ActualClosureDate?: string;
  
  // Common fields
  Effectiveness?: string;
  Remarks?: string;
  
  // SharePoint metadata
  Created?: string;
  Modified?: string;
  AuthorId?: number;
  EditorId?: number;
}

/**
 * RAID List Service
 * Specialized service for RAID (Risk, Assumption, Issue, Dependency, Opportunity, Constraints) items
 * Uses the generic SharePoint service underneath
 */
export class RaidListService {
  private genericService: IGenericService;
  private context: WebPartContext;
  private listName: string;
  private enablePeoplePickerFields: boolean = true;

  constructor(context: WebPartContext, listName: string = LIST_NAMES.RAID_LOGS) {
    this.context = context;
    this.genericService = GenericServiceInstance;
    this.genericService.init(undefined, context);
    this.listName = listName;
  }

  private createQueryOptions(additionalOptions?: Partial<IListQueryOptions>): IListQueryOptions {
    const baseOptions: IListQueryOptions = {};
    
    if (this.enablePeoplePickerFields) {
      baseOptions.expand = ['Responsibility', 'ByWhom'];
      baseOptions.select = [
        '*',
        'Responsibility/Id',
        'Responsibility/Title',
        'ResponsibilityId/Email',
        'ResponsibilityId/LoginName',
        'ByWhom/Id',
        'ByWhom/Title',
        'ByWhomId/Email',
        'ByWhomId/LoginName'
      ];
    } else {
      baseOptions.select = ['*'];
    }
    
    return { ...baseOptions, ...additionalOptions };
  }

  private async convertToSharePointItem(raidItem: IRaidItem): Promise<Omit<IRaidSharePointItem, 'Id'>> {
    const spItem: any = {
      SelectType: raidItem.type,
      RAIDId: raidItem.raidId || undefined,
      IdentificationDate: this.convertToISODate(raidItem.identificationDate),
      RiskDescription: raidItem.description || undefined,
      AssociatedGoal: raidItem.associatedGoal || undefined,
      RiskSource: raidItem.source || undefined,
      RiskCategory: raidItem.category || undefined,
      Impact: raidItem.impact || undefined,
      RiskPriority: raidItem.priority || undefined,
      ImpactValue: this.convertToNumber(raidItem.impactValue),
      ProbabilityValue: this.convertToNumber(raidItem.probabilityValue),
      RiskExposure: this.convertToNumber(raidItem.riskExposure),
      TypeOfAction: raidItem.typeOfAction || undefined,
      PotentialCost: this.convertToNumber(raidItem.potentialCost),
      PotentialBenefit: this.convertToNumber(raidItem.potentialBenefit),
      OpportunityValue: this.convertToNumberWithDecimals(raidItem.opportunityValue, 2),
      ActionPlan: raidItem.actionPlan || undefined,
      TargetDate: this.convertToISODate(raidItem.targetDate),
      ActualDate: this.convertToISODate(raidItem.actualDate),
      RiskStatus: raidItem.status || undefined,
      IssueDetails: raidItem.details || undefined,
      IDADate: this.convertToISODate(raidItem.date),
      ImplementationActions: raidItem.implementationActions || undefined,
      PlannedClosureDate: this.convertToISODate(raidItem.plannedClosureDate),
      ActualClosureDate: this.convertToISODate(raidItem.actualClosureDate),
      Effectiveness: raidItem.effectiveness || undefined,
      Remarks: raidItem.remarks || undefined
      // DO NOT set responsibility or byWhom here - they will be converted to ResponsibilityId and ByWhomId below
    };

    if (this.enablePeoplePickerFields) {
      try {
        if (raidItem.responsibility) {
          console.log('🔄 Converting responsibility field:', raidItem.responsibility);
          const responsibilityValue = await this.convertUserFieldForSharePointAsync(raidItem.responsibility, this.context);
          console.log('✅ Converted responsibility to:', responsibilityValue);
          if (responsibilityValue !== null && responsibilityValue !== undefined) {
            // Always use Id suffix for people picker fields
            spItem.ResponsibilityId = responsibilityValue;
          }
        }
      } catch (error) {
        console.error('Error converting responsibility field:', error);
        this.enablePeoplePickerFields = false;
      }

      try {
        if (raidItem.byWhom) {
          console.log('🔄 Converting byWhom field:', raidItem.byWhom);
          const byWhomValue = await this.convertUserFieldForSharePointAsync(raidItem.byWhom, this.context);
          console.log('✅ Converted byWhom to:', byWhomValue);
          if (byWhomValue !== null && byWhomValue !== undefined) {
            // Always use Id suffix for people picker fields
            spItem.ByWhomId = byWhomValue;
          }
        }
      } catch (error) {
        console.error('Error converting byWhom field:', error);
        this.enablePeoplePickerFields = false;
      }
    }

    Object.keys(spItem).forEach(key => {
      if (spItem[key as keyof typeof spItem] === undefined) {
        delete spItem[key as keyof typeof spItem];
      }
    });

    return spItem;
  }

  private convertFromSharePointItem(spItem: IRaidSharePointItem): IExtendedRaidItem {
    const raidItem: IExtendedRaidItem = {
      id: spItem.Id || 0,
      type: spItem.SelectType,
      raidId: spItem.RAIDId,
      identificationDate: spItem.IdentificationDate,
      description: spItem.RiskDescription,
      associatedGoal: spItem.AssociatedGoal,
      source: spItem.RiskSource,
      category: spItem.RiskCategory,
      impact: spItem.Impact,
      priority: spItem.RiskPriority,
      impactValue: spItem.ImpactValue,
      probabilityValue: spItem.ProbabilityValue,
      riskExposure: spItem.RiskExposure,
      typeOfAction: spItem.TypeOfAction,
      potentialCost: spItem.PotentialCost,
      potentialBenefit: spItem.PotentialBenefit,
      opportunityValue: spItem.OpportunityValue,
      actionPlan: spItem.ActionPlan,
      responsibility: this.parsePeoplePickerField((spItem as any).Responsibility || spItem.ResponsibilityId?.[0]),
      targetDate: spItem.TargetDate,
      actualDate: spItem.ActualDate,
      status: spItem.RiskStatus,
      details: spItem.IssueDetails,
      date: spItem.IDADate,
      byWhom: this.parsePeoplePickerField((spItem as any).ByWhom || spItem.ByWhomId?.[0]),
      implementationActions: spItem.ImplementationActions,
      plannedClosureDate: spItem.PlannedClosureDate,
      actualClosureDate: spItem.ActualClosureDate,
      effectiveness: spItem.Effectiveness,
      remarks: spItem.Remarks
    };

    return raidItem;
  }

  private parsePeoplePickerField(value: any): any {
    if (!value) return undefined;
    
    try {
      if (typeof value === 'object') {
        if (Array.isArray(value)) {
          const parsedUsers = value.map((user: any) => this.parseUserObject(user)).filter((user: any) => user !== null && user !== undefined);
          return parsedUsers.length > 0 ? parsedUsers : undefined;
        }
        
        if (value.results && Array.isArray(value.results)) {
          const parsedUsers = value.results.map((user: any) => this.parseUserObject(user)).filter((user: any) => user !== null && user !== undefined);
          return parsedUsers.length > 0 ? parsedUsers : undefined;
        }
        
        const parsedUser = this.parseUserObject(value);
        if (parsedUser) {
          return [parsedUser];
        }
      }
      
      if (typeof value === 'string') {
        try {
          const parsed = JSON.parse(value);
          return this.parsePeoplePickerField(parsed);
        } catch {
          if (value.indexOf('@') !== -1) {
            return [{
              id: value,
              email: value,
              loginName: value,
              displayName: value.split('@')[0]
            }];
          } else {
            const numericValue = parseInt(value, 10);
            if (!isNaN(numericValue) && numericValue > 0) {
              return [{
                id: String(numericValue),
                loginName: '',
                displayName: `User ${numericValue}`,
                email: ''
              }];
            }
            return undefined;
          }
        }
      }
      
      if (typeof value === 'number') {
        return [{
          id: String(value),
          loginName: '',
          displayName: `User ${value}`,
          email: ''
        }];
      }
      
      return undefined;
    } catch (error) {
      return undefined;
    }
  }

  private parseUserObject(user: any): any {
    if (!user || typeof user !== 'object') return user;
    
    try {
      const parsedUser: any = {};
      
      if (user.Id !== undefined) parsedUser.id = String(user.Id);
      else if (user.id !== undefined) parsedUser.id = String(user.id);
      else if (user.ID !== undefined) parsedUser.id = String(user.ID);
      
      if (user.Email) parsedUser.email = user.Email;
      else if (user.email) parsedUser.email = user.email;
      else if (user.EMail) parsedUser.email = user.EMail;
      
      if (user.LoginName) parsedUser.loginName = user.LoginName;
      else if (user.loginName) parsedUser.loginName = user.loginName;
      else if (user.UserName) parsedUser.loginName = user.UserName;
      else if (user.userName) parsedUser.loginName = user.userName;
      
      if (user.Title) parsedUser.displayName = user.Title;
      else if (user.DisplayName) parsedUser.displayName = user.DisplayName;
      else if (user.displayName) parsedUser.displayName = user.displayName;
      else if (user.text) parsedUser.displayName = user.text;
      else if (user.name) parsedUser.displayName = user.name;
      else if (user.Name) parsedUser.displayName = user.Name;
      else if (user.title) parsedUser.displayName = user.title;
      
      if (!parsedUser.id) {
        if (parsedUser.loginName) parsedUser.id = parsedUser.loginName;
        else if (parsedUser.email) parsedUser.id = parsedUser.email;
      }
      
      if (!parsedUser.loginName && parsedUser.email) {
        parsedUser.loginName = parsedUser.email;
      }
      if (!parsedUser.email && parsedUser.loginName && parsedUser.loginName.indexOf('@') !== -1) {
        parsedUser.email = parsedUser.loginName;
      }
      if (!parsedUser.displayName && parsedUser.email) {
        parsedUser.displayName = parsedUser.email.split('@')[0];
      }
      
      if (!parsedUser.id && !parsedUser.email && !parsedUser.loginName) {
        return user;
      }
      
      return parsedUser;
    } catch (error) {
      return user;
    }
  }

  private async convertUserFieldForSharePointAsync(userField: any, context: WebPartContext): Promise<number[] | null> {
    if (!userField) return null;
    
    try {
      // Parse the serialized format "id|email; id|email" (same as RCA repository)
      // The PeoplePicker with ensureUser={true} should provide numeric IDs
      if (typeof userField === 'string') {
        // Empty string check
        if (userField.trim() === '') return null;
        
        const parts = userField.split(/;\s*/);
        const userIds: number[] = [];
        
        for (const part of parts) {
          if (!part || part.trim() === '') continue;
          
          // Split by pipe to get ID|email format
          const pipeIndex = part.indexOf('|');
          const idPart = pipeIndex !== -1 ? part.substring(0, pipeIndex).trim() : part.trim();
          
          // Parse the ID as number (PeoplePicker with ensureUser should give numeric IDs)
          const numericId = parseInt(idPart, 10);
          if (!isNaN(numericId) && numericId > 0) {
            userIds.push(numericId);
          } else {
            console.warn('⚠️ Non-numeric ID found in serialized user field:', idPart, '- Ensure ensureUser={true} is set on PeoplePicker');
          }
        }
        
        console.log('📊 Parsed user IDs for SharePoint:', userIds);
        
        // Return array of IDs (PnPjs will handle the proper formatting)
        return userIds.length > 0 ? userIds : null;
      }
      
      // Handle array format
      if (Array.isArray(userField)) {
        const userIds: number[] = [];
        
        for (const user of userField) {
          if (typeof user === 'number' && !isNaN(user) && user > 0) {
            userIds.push(user);
          } else if (typeof user === 'string' && user.trim() !== '') {
            const pipeIndex = user.indexOf('|');
            const idPart = pipeIndex !== -1 ? user.substring(0, pipeIndex).trim() : user.trim();
            
            const numericId = parseInt(idPart, 10);
            if (!isNaN(numericId) && numericId > 0) {
              userIds.push(numericId);
            }
          }
        }
        
        // Return array of IDs (PnPjs will handle the proper formatting)
        return userIds.length > 0 ? userIds : null;
      }
      
      return null;
    } catch (error) {
      console.error('❌ Error in convertUserFieldForSharePointAsync:', error);
      return null;
    }
  }

  /**
   * Helper method to convert values to numbers for SharePoint numeric fields
   */
  private convertToNumber(value: any): number | undefined {
    // Handle null, undefined, empty string, and whitespace-only strings
    if (value === null || value === undefined || value === '' || 
        (typeof value === 'string' && value.trim() === '')) {
      return undefined;
    }
    
    // If it's already a valid number, return it
    if (typeof value === 'number' && !isNaN(value) && isFinite(value)) {
      return value;
    }
    
    // If it's a string, try to parse it
    if (typeof value === 'string') {
      const trimmed = value.trim();
      // Skip empty strings after trimming
      if (trimmed === '') {
        return undefined;
      }
      
      const parsed = parseFloat(trimmed);
      return (isNaN(parsed) || !isFinite(parsed)) ? undefined : parsed;
    }
    
    // For boolean values, convert true=1, false=0
    if (typeof value === 'boolean') {
      return value ? 1 : 0;
    }
    
    // For any other type, try to convert to number
    const converted = Number(value);
    return (isNaN(converted) || !isFinite(converted)) ? undefined : converted;
  }

  private convertToNumberWithDecimals(value: any, decimals: number = 2): number | undefined {
    const numValue = this.convertToNumber(value);
    if (numValue === undefined) {
      return undefined;
    }
    // Round to specified decimal places
    return Math.round(numValue * Math.pow(10, decimals)) / Math.pow(10, decimals);
  }

  /**
   * Helper method to convert dates to ISO string format for SharePoint
   */
  private convertToISODate(value: any): string | undefined {
    if (value === null || value === undefined || value === '') {
      return undefined;
    }
    
    try {
      // If it's already a valid Date object
      if (value instanceof Date && !isNaN(value.getTime())) {
        return value.toISOString();
      }
      
      // If it's a string, try to parse it
      if (typeof value === 'string') {
        const date = new Date(value);
        if (!isNaN(date.getTime())) {
          return date.toISOString();
        }
      }
      
      if (typeof value === 'number' && !isNaN(value)) {
        const date = new Date(value);
        if (!isNaN(date.getTime())) {
          return date.toISOString();
        }
      }
      
      return undefined;
    } catch (error) {
      return undefined;
    }
  }

  /**
   * CREATE: Add new RAID item
   */
  async createRaidItem(raidItem: Omit<IRaidItem, 'id'>): Promise<IExtendedRaidItem | null> {
    try {
      const spItem = await this.convertToSharePointItem({ ...raidItem, id: 0 });
      // Clean the item using RaidLogs-specific cleaning function
      const cleanedItem = this.genericService.cleanItemForRaidSave(spItem);
      const queryOptions = this.createQueryOptions();

      const result = await this.genericService.saveItem<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        item: cleanedItem,
        select: queryOptions.select || [],
        expand: queryOptions.expand || []
      });

      if (result && result.success) {
        let createdSpItem: IRaidSharePointItem | undefined = undefined;
        if (result.item) {
          createdSpItem = result.item as IRaidSharePointItem;
        } else if (result.itemId) {
          const fetched = await this.genericService.fetchAllItems<IRaidSharePointItem>({
            context: this.context,
            listTitle: this.listName,
            filter: `Id eq ${result.itemId}`,
            select: queryOptions.select || [],
            expand: queryOptions.expand || []
          });
          createdSpItem = fetched && fetched.length > 0 ? fetched[0] : undefined;
        }

        if (createdSpItem) {
          return this.convertFromSharePointItem(createdSpItem);
        }
      }

      return null;
    } catch (error) {
      return null;
    }
  }

  /**
   * CREATE: Create Risk item with separate Mitigation and Contingency SharePoint items
   * This method creates 2 SharePoint list items for a single Risk form entry:
   * - One item for Mitigation action
   * - One item for Contingency action
   * Both items share the same RaidID and common field values
   * Only action-specific fields differ between the two items
   */
  async createRiskItemWithActions(
    raidItem: Omit<IRaidItem, 'id'>,
    mitigationAction: IRaidAction | null,
    contingencyAction: IRaidAction | null
  ): Promise<IExtendedRaidItem[] | null> {
    try {
      const createdItems: IExtendedRaidItem[] = [];

      if (!raidItem.raidId) {
        return null;
      }

      if (mitigationAction) {
        const mitigationItem: Omit<IRaidItem, 'id'> = {
          ...raidItem,
          typeOfAction: 'Mitigation',
          actionPlan: mitigationAction.plan,
          responsibility: mitigationAction.responsibility,
          targetDate: mitigationAction.targetDate,
          actualDate: mitigationAction.actualDate,
          status: mitigationAction.status
        };

        const createdMitigation = await this.createRaidItem(mitigationItem);
        if (createdMitigation) {
          createdItems.push(createdMitigation);
        } else {
          return null;
        }
      }

      if (contingencyAction) {
        const contingencyItem: Omit<IRaidItem, 'id'> = {
          ...raidItem,
          typeOfAction: 'Contingency',
          actionPlan: contingencyAction.plan,
          responsibility: contingencyAction.responsibility,
          targetDate: contingencyAction.targetDate,
          actualDate: contingencyAction.actualDate,
          status: contingencyAction.status
        };

        const createdContingency = await this.createRaidItem(contingencyItem);
        if (createdContingency) {
          createdItems.push(createdContingency);
        } else {
          return null;
        }
      }

      return createdItems.length > 0 ? createdItems : null;
    } catch (error) {
      return null;
    }
  }

  async getAllRaidItems(): Promise<IExtendedRaidItem[]> {
    try {
      const options = this.createQueryOptions({
        orderBy: 'Modified'
      });

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        orderBy: options.orderBy,
        filter: options.filter
      });

      return items.map((spItem: IRaidSharePointItem) => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  async getRaidItemsByType(type: RaidType): Promise<IExtendedRaidItem[]> {
    try {
      const options = this.createQueryOptions({
        filter: `SelectType eq '${type}'`,
        orderBy: 'Modified'
      });

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: options.filter,
        orderBy: options.orderBy
      });

      return items.map(spItem => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  async getRaidItemById(itemId: number): Promise<IExtendedRaidItem | null> {
    try {
      const options = this.createQueryOptions();

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: `Id eq ${itemId}`
      });

      if (items && items.length > 0) {
        return this.convertFromSharePointItem(items[0]);
      }
      return null;
    } catch (error) {
      return null;
    }
  }

  async updateRaidItem(itemId: number, updates: Partial<IRaidItem>): Promise<IExtendedRaidItem | null> {
    try {
      const currentItem = await this.getRaidItemById(itemId);
      if (!currentItem) {
        return null;
      }

      const updatedItem = { ...currentItem, ...updates };
      const spUpdates = await this.convertToSharePointItem(updatedItem);
      // Clean the item using RaidLogs-specific cleaning function
      const cleanedUpdates = this.genericService.cleanItemForRaidSave(spUpdates);

      const queryOptions = this.createQueryOptions();

      const result = await this.genericService.updateItem<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        itemId,
        item: cleanedUpdates,
        select: queryOptions.select || [],
        expand: queryOptions.expand || []
      });

      if (result && result.success) {
        let updatedSpItem: IRaidSharePointItem | undefined = undefined;
        if (result.item) {
          updatedSpItem = result.item as IRaidSharePointItem;
        } else if (result.itemId) {
          const fetched = await this.genericService.fetchAllItems<IRaidSharePointItem>({
            context: this.context,
            listTitle: this.listName,
            filter: `Id eq ${result.itemId}`,
            select: queryOptions.select || [],
            expand: queryOptions.expand || []
          });
          updatedSpItem = fetched && fetched.length > 0 ? fetched[0] : undefined;
        }

        if (updatedSpItem) {
          return this.convertFromSharePointItem(updatedSpItem);
        }
      }
      return null;
    } catch (error) {
      return null;
    }
  }

  async deleteRaidItem(itemId: number): Promise<boolean> {
    try {
      const result = await this.genericService.deleteItem({
        context: this.context,
        listTitle: this.listName,
        itemId
      });

      return result && result.success;
    } catch (error) {
      return false;
    }
  }

  /**
   * UTILITY: Get Risk items by RaidID
   * Returns all SharePoint items that share the same RaidID
   * Used to fetch both Mitigation and Contingency items for a single Risk
   */
  async getRiskItemsByRaidId(raidId: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `RAIDId eq '${raidId}'`,
        orderBy: 'TypeOfAction'
      };

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: options.filter,
        orderBy: options.orderBy
      });

      return items.map((spItem: IRaidSharePointItem) => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  async updateRiskItemsByRaidId(
    raidId: string,
    updates: Partial<IRaidItem>,
    mitigationAction: IRaidAction | null,
    contingencyAction: IRaidAction | null
  ): Promise<boolean> {
    try {
      const existingItems = await this.getRiskItemsByRaidId(raidId);
      
      if (existingItems.length === 0) {
        return false;
      }

      let existingMitigation: IExtendedRaidItem | undefined;
      let existingContingency: IExtendedRaidItem | undefined;
      
      for (let i = 0; i < existingItems.length; i++) {
        if (existingItems[i].typeOfAction === 'Mitigation') {
          existingMitigation = existingItems[i];
        } else if (existingItems[i].typeOfAction === 'Contingency') {
          existingContingency = existingItems[i];
        }
      }

      let allSuccess = true;

      if (mitigationAction) {
        if (existingMitigation) {
          const itemUpdates: Partial<IRaidItem> = {
            ...updates,
            typeOfAction: 'Mitigation',
            actionPlan: mitigationAction.plan,
            responsibility: mitigationAction.responsibility,
            targetDate: mitigationAction.targetDate,
            actualDate: mitigationAction.actualDate,
            status: mitigationAction.status
          };
          
          const success = await this.updateRaidItem(existingMitigation.id, itemUpdates);
          if (!success) {
            allSuccess = false;
          }
        } else {
          const mitigationItem: Omit<IRaidItem, 'id'> = {
            ...updates,
            raidId,
            type: 'Risk',
            typeOfAction: 'Mitigation',
            actionPlan: mitigationAction.plan,
            responsibility: mitigationAction.responsibility,
            targetDate: mitigationAction.targetDate,
            actualDate: mitigationAction.actualDate,
            status: mitigationAction.status
          };
          
          const created = await this.createRaidItem(mitigationItem);
          if (!created) {
            allSuccess = false;
          }
        }
      } else if (existingMitigation) {
        const success = await this.deleteRaidItem(existingMitigation.id);
        if (!success) {
          allSuccess = false;
        }
      }

      if (contingencyAction) {
        if (existingContingency) {
          const itemUpdates: Partial<IRaidItem> = {
            ...updates,
            typeOfAction: 'Contingency',
            actionPlan: contingencyAction.plan,
            responsibility: contingencyAction.responsibility,
            targetDate: contingencyAction.targetDate,
            actualDate: contingencyAction.actualDate,
            status: contingencyAction.status
          };
          
          const success = await this.updateRaidItem(existingContingency.id, itemUpdates);
          if (!success) {
            allSuccess = false;
          }
        } else {
          const contingencyItem: Omit<IRaidItem, 'id'> = {
            ...updates,
            raidId,
            type: 'Risk',
            typeOfAction: 'Contingency',
            actionPlan: contingencyAction.plan,
            responsibility: contingencyAction.responsibility,
            targetDate: contingencyAction.targetDate,
            actualDate: contingencyAction.actualDate,
            status: contingencyAction.status
          };
          
          const created = await this.createRaidItem(contingencyItem);
          if (!created) {
            allSuccess = false;
          }
        }
      } else if (existingContingency) {
        const success = await this.deleteRaidItem(existingContingency.id);
        if (!success) {
          allSuccess = false;
        }
      }

      return allSuccess;
    } catch (error) {
      return false;
    }
  }

  async deleteRiskItemsByRaidId(raidId: string): Promise<boolean> {
    try {
      const itemsToDelete = await this.getRiskItemsByRaidId(raidId);
      
      if (itemsToDelete.length === 0) {
        return false;
      }

      let allSuccess = true;

      for (const item of itemsToDelete) {
        const success = await this.deleteRaidItem(item.id);
        if (!success) {
          allSuccess = false;
        }
      }

      return allSuccess;
    } catch (error) {
      return false;
    }
  }

  async getRaidItemsByStatus(status: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `RiskStatus eq '${status}'`,
        orderBy: 'Modified'
      };

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: options.filter,
        orderBy: options.orderBy
      });

      return items.map((spItem: IRaidSharePointItem) => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  async getRaidItemsByPriority(priority: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `RiskPriority eq '${priority}'`,
        orderBy: 'Modified'
      };

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: options.filter,
        orderBy: options.orderBy
      });

      return items.map((spItem: IRaidSharePointItem) => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  async searchRaidItems(searchTerm: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `substringof('${searchTerm}', Title) or substringof('${searchTerm}', RiskDescription) or substringof('${searchTerm}', IssueDetails)`,
        orderBy: 'Modified'
      };

      const items = await this.genericService.fetchAllItems<IRaidSharePointItem>({
        context: this.context,
        listTitle: this.listName,
        select: options.select || [],
        expand: options.expand || [],
        filter: options.filter,
        orderBy: options.orderBy
      });

      return items.map((spItem: IRaidSharePointItem) => this.convertFromSharePointItem(spItem));
    } catch (error) {
      return [];
    }
  }

  /**
   * UTILITY: Get high priority items
   */
  async getHighPriorityItems(): Promise<IExtendedRaidItem[]> {
    return await this.getRaidItemsByPriority('High');
  }

  /**
   * UTILITY: Get open items
   */
  async getOpenItems(): Promise<IExtendedRaidItem[]> {
    return await this.getRaidItemsByStatus('Open');
  }

  async bulkUpdateStatus(itemIds: number[], newStatus: string): Promise<boolean> {
    try {
      let allSuccess = true;

      for (const itemId of itemIds) {
        const success = await this.updateRaidItem(itemId, { status: newStatus });
        if (!success) {
          allSuccess = false;
        }
      }

      return allSuccess;
    } catch (error) {
      return false;
    }
  }

  /**
   * Get version history for a specific list item
   */
  async getVersionHistory(itemId: number): Promise<any[]> {
    try {
      if (!itemId || itemId <= 0) return [];

      const versions = await this.genericService.getVersionHistory<any>({
        context: this.context,
        listTitle: this.listName,
        itemId,
        // Request fields relevant to RAID log items with Editor expanded
        select: [
          '*',
          'Editor/Title', 
          'Editor/EMail',
          'Editor/Name'
        ],
        expand: ['Editor']
      });

      return versions || [];
    } catch (error) {
      console.error('Error fetching version history:', error);
      return [];
    }
  }

  /**
   * UTILITY: Check if RAID list exists
   */
  async checkListExists(): Promise<boolean> {
    try {
      await this.genericService.fetchAllItems({ context: this.context, listTitle: this.listName, pageSize: 1 });
      return true;
    } catch (error) {
      return false;
    }
  }

  /**
   * UTILITY: Get list information
   */
  async getListInfo(): Promise<any> {
    try {
      // GenericService does not currently expose list metadata; try a minimal fetch to confirm availability
      await this.genericService.fetchAllItems({ context: this.context, listTitle: this.listName, pageSize: 1 });
      return { listName: this.listName };
    } catch (error) {
      return null;
    }
  }

  /**
   * DROPDOWN OPTIONS: Fetch dropdown options from SharePoint lists
   * These methods fetch dropdown options from SharePoint lists dynamically
   * Title field is used as key and Text field is used as text
   */

  /**
   * Fetch POTENTIAL_COST options from SharePoint list
   * @returns Array of dropdown options with key (Title) and text (Text)
   */
  async getPotentialCostOptions(): Promise<Array<{ key: string; text: string }>> {
    try {
      const items = await this.genericService.fetchAllItems<any>({
        context: this.context,
        listTitle: LIST_NAMES.POTENTIAL_COST,
        select: ['Title', 'Text']
      });

      // Sort by numeric value of Title (1-10)
      const sortedItems = items.sort((a, b) => Number(a.Title) - Number(b.Title));

      return sortedItems.map(item => ({
        key: item.Title,
        text: `${item.Title} - ${item.Text}`
      }));
    } catch (error) {
      console.error('Error fetching POTENTIAL_COST options:', error);
      return [];
    }
  }

  /**
   * Fetch POTENTIAL_BENEFIT options from SharePoint list
   * @returns Array of dropdown options with key (Title) and text (Text)
   */
  async getPotentialBenefitOptions(): Promise<Array<{ key: string; text: string }>> {
    try {
      const items = await this.genericService.fetchAllItems<any>({
        context: this.context,
        listTitle: LIST_NAMES.POTENTIAL_BENEFIT,
        select: ['Title', 'Text']
      });

      // Sort by numeric value of Title (1-10)
      const sortedItems = items.sort((a, b) => Number(a.Title) - Number(b.Title));

      return sortedItems.map(item => ({
        key: item.Title,
        text: `${item.Title} - ${item.Text}`
      }));
    } catch (error) {
      console.error('Error fetching POTENTIAL_BENEFIT options:', error);
      return [];
    }
  }

  /**
   * Fetch PROBABILITY_VALUE options from SharePoint list
   * @returns Array of dropdown options with key (Title) and text (Text)
   */
  async getProbabilityValueOptions(): Promise<Array<{ key: string; text: string }>> {
    try {
      const items = await this.genericService.fetchAllItems<any>({
        context: this.context,
        listTitle: LIST_NAMES.PROBABILITY_VALUE,
        select: ['Title', 'Text']
      });

      // Sort by numeric value of Title (1-10)
      const sortedItems = items.sort((a, b) => Number(a.Title) - Number(b.Title));

      return sortedItems.map(item => ({
        key: item.Title,
        text: `${item.Title} - ${item.Text}`
      }));
    } catch (error) {
      console.error('Error fetching PROBABILITY_VALUE options:', error);
      return [];
    }
  }

  /**
   * Fetch IMPACT_VALUE options from SharePoint list
   * @returns Array of dropdown options with key (Title) and text (Text)
   */
  async getImpactValueOptions(): Promise<Array<{ key: string; text: string }>> {
    try {
      const items = await this.genericService.fetchAllItems<any>({
        context: this.context,
        listTitle: LIST_NAMES.IMPACT_VALUE,
        select: ['Title', 'Text']
      });

      // Sort by numeric value of Title (1-10)
      const sortedItems = items.sort((a, b) => Number(a.Title) - Number(b.Title));

      return sortedItems.map(item => ({
        key: item.Title,
        text: `${item.Title} - ${item.Text}`
      }));
    } catch (error) {
      console.error('Error fetching IMPACT_VALUE options:', error);
      return [];
    }
  }

}

/**
 * RAID Service Factory
 * Creates a singleton instance of RAID service
 */
export class RaidServiceFactory {
  private static instance: RaidListService;

  static getInstance(context: WebPartContext, listName?: string): RaidListService {
    if (!RaidServiceFactory.instance) {
      RaidServiceFactory.instance = new RaidListService(context, listName || LIST_NAMES.RAID_LOGS);
    }
    return RaidServiceFactory.instance;
  }

  /**
   * Reset the singleton instance (useful for testing or context changes)
   */
  static resetInstance(): void {
    RaidServiceFactory.instance = undefined as any;
  }
}