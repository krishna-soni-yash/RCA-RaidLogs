import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService, ISharePointListItem, IListQueryOptions } from './SharePointService';
import { IRaidItem, RaidType } from '../components/RaidLogs/IRaidItem';
import { LIST_NAMES, ERROR_MESSAGES } from '../constants/Constants';
import { 
  IExtendedRaidItem
} from '../interfaces/IRaidService';

/**
 * SharePoint List Item interface for RAID items
 * Extends the generic interface with RAID-specific fields
 * Field names match SharePoint internal field names
 */
export interface IRaidSharePointItem extends ISharePointListItem {
  Id?: number;
  Title?: string;
  SelectType: RaidType; // Internal name for RAID Type selection
  IdentificationDate?: string;
  RiskDescription?: string; // Internal name for Description
  Applicability?: string;
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
  ResponsibilityId?: any; // People picker field with Id suffix (used for saving/updating)
  TargetDate?: string;
  ActualDate?: string;
  RiskStatus?: string; // Internal name for Status
  
  // Issue/Assumption/Dependency specific fields
  IssueDetails?: string; // Internal name for Details
  IDADate?: string; // Internal name for Date
  ByWhom?: any; // Expanded people picker field (when using expand in query)
  ByWhomId?: any; // People picker field with Id suffix (used for saving/updating)
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
 * Specialized service for RAID (Risk, Assumption, Issue, Dependency) items
 * Uses the generic SharePoint service underneath
 */
export class RaidListService {
  private spService: SharePointService;
  private listName: string;
  private enablePeoplePickerFields: boolean = true; // Feature flag for people picker fields - ENABLED

  constructor(context: WebPartContext, listName: string = LIST_NAMES.RAID_LOGS) {
    this.spService = new SharePointService(context);
    this.listName = listName;
  }

  /**
   * Manually disable people picker fields if they're causing issues
   */
  public disablePeoplePickerFields(): void {
    this.enablePeoplePickerFields = false;
    console.log('People picker fields manually disabled');
  }

  /**
   * Re-enable people picker fields
   */
  public enablePeoplePickerFieldsManually(): void {
    this.enablePeoplePickerFields = true;
    console.log('People picker fields manually enabled');
  }

  /**
   * Test method for people picker field conversion
   */
  public testPeoplePickerConversion(testData: any): any {
    console.log('🧪 Testing people picker conversion with data:', testData);
    const result = this.convertUserFieldForSharePoint(testData);
    console.log('🧪 Conversion result:', result);
    return result;
  }

  /**
   * Validate that byWhom and responsibility fields are properly converted to user IDs
   * This method helps ensure both fields save user IDs when possible
   */
  public validatePeoplePickerData(raidItem: any): { responsibility: any; byWhom: any; warnings: string[] } {
    const warnings: string[] = [];
    
    console.log('🔍 Validating people picker data for byWhom and responsibility fields...');
    
    // Validate responsibility field
    let responsibilityResult = null;
    if (raidItem.responsibility) {
      responsibilityResult = this.convertUserFieldForSharePoint(raidItem.responsibility);
      if (typeof responsibilityResult === 'string' && responsibilityResult.indexOf('@') === -1) {
        warnings.push('Responsibility field could not be converted to user ID, using string value');
      } else if (typeof responsibilityResult === 'number') {
        console.log('✅ Responsibility field successfully converted to user ID:', responsibilityResult);
      }
    }
    
    // Validate byWhom field  
    let byWhomResult = null;
    if (raidItem.byWhom) {
      byWhomResult = this.convertUserFieldForSharePoint(raidItem.byWhom);
      if (typeof byWhomResult === 'string' && byWhomResult.indexOf('@') === -1) {
        warnings.push('ByWhom field could not be converted to user ID, using string value');
      } else if (typeof byWhomResult === 'number') {
        console.log('✅ ByWhom field successfully converted to user ID:', byWhomResult);
      }
    }
    
    return {
      responsibility: responsibilityResult,
      byWhom: byWhomResult,
      warnings
    };
  }

  /**
   * Comprehensive test method for people picker functionality
   * Tests both saving and fetching of people picker data
   */
  public async testPeoplePickerFunctionality(testUserId?: number, testEmail?: string): Promise<void> {
    console.log('🧪 Starting comprehensive people picker functionality test...');
    
    try {
      // Enable people picker fields for testing
      this.enablePeoplePickerFieldsManually();
      
      // Test data conversion - focusing on user ID preference for byWhom and responsibility
      const testCases = [
        { name: 'User ID (preferred)', data: testUserId || 11 },
        { name: 'User ID as String', data: String(testUserId || 11) },
        { name: 'User Object with ID', data: { id: testUserId || 11, email: testEmail || 'test@example.com', title: 'Test User' } },
        { name: 'User Object with Email fallback', data: { email: testEmail || 'test@example.com', title: 'Test User' } },
        { name: 'Login Name', data: `i:0#.f|membership|${testEmail || 'test@example.com'}` },
        { name: 'Array of Mixed Users', data: [testUserId || 11, { id: 12, email: 'test2@example.com' }, testEmail || 'test3@example.com'] }
      ];
      
      console.log('🧪 Testing data conversion for different formats:');
      testCases.forEach(testCase => {
        console.log(`\n📝 Testing ${testCase.name}:`);
        console.log('Input:', testCase.data);
        const converted = this.convertUserFieldForSharePoint(testCase.data);
        console.log('Converted:', converted);
      });
      
      // Test parsing functionality
      const mockSharePointResponses = [
        { name: 'Single User Object', data: { Id: 11, Title: 'Test User', Email: 'test@example.com', LoginName: 'i:0#.f|membership|test@example.com' } },
        { name: 'User Array', data: [{ Id: 11, Title: 'Test User', Email: 'test@example.com' }, { Id: 12, Title: 'Test User 2', Email: 'test2@example.com' }] },
        { name: 'OData Results', data: { results: [{ Id: 11, Title: 'Test User', Email: 'test@example.com' }] } }
      ];
      
      console.log('\n🧪 Testing response parsing for different SharePoint formats:');
      mockSharePointResponses.forEach(testCase => {
        console.log(`\n📝 Testing ${testCase.name}:`);
        console.log('SharePoint Response:', testCase.data);
        const parsed = this.parsePeoplePickerField(testCase.data);
        console.log('Parsed Result:', parsed);
      });
      
      // Test specific validation for byWhom and responsibility fields
      console.log('\n🧪 Testing byWhom and responsibility field validation:');
      const mockRaidItem = {
        responsibility: testUserId || 11,
        byWhom: testEmail || 'test@example.com'
      };
      
      const validation = this.validatePeoplePickerData(mockRaidItem);
      console.log('Validation Result:', validation);
      
      if (validation.warnings.length > 0) {
        console.warn('⚠️ Validation warnings:', validation.warnings);
      } else {
        console.log('✅ All people picker fields validated successfully');
      }
      
      console.log('\n✅ People picker functionality test completed');
    } catch (error) {
      console.error('❌ Error during people picker functionality test:', error);
    }
  }

  /**
   * Test query with people picker fields
   * Use this to test actual SharePoint queries with people picker expansion
   */
  public async testPeoplePickerQuery(itemId?: number): Promise<any> {
    console.log('🧪 Testing people picker query...');
    
    try {
      // Enable people picker fields
      this.enablePeoplePickerFieldsManually();
      
      const testOptions = this.createPeoplePickerTestQueryOptions();
      console.log('Test query options:', testOptions);
      
      if (itemId) {
        // Test single item query
        const result = await this.spService.getItemById<IRaidSharePointItem>(
          { listName: this.listName },
          itemId,
          testOptions
        );
        
        console.log('Single item query result:', result);
        
        if (result.success && result.data) {
          console.log('ResponsibilityId field:', result.data.Responsibility);
          console.log('ByWhomId field:', result.data.ByWhom);
          
          // Test parsing
          const parsedResponsibility = this.parsePeoplePickerField(result.data.Responsibility);
          const parsedByWhom = this.parsePeoplePickerField(result.data.ByWhom);
          
          console.log('Parsed Responsibility:', parsedResponsibility);
          console.log('Parsed ByWhom:', parsedByWhom);
        }
        
        return result;
      } else {
        // Test list query
        const listOptions = this.createPeoplePickerTestQueryOptions({ top: 5 });
        const result = await this.spService.getItems<IRaidSharePointItem>(
          { listName: this.listName },
          listOptions
        );
        
        console.log('List query result:', result);
        
        if (result.success && result.data && result.data.length > 0) {
          result.data.forEach((item, index) => {
            console.log(`Item ${index + 1}:`);
            console.log('  Responsibility:', item.Responsibility);
            console.log('  ByWhom:', item.ByWhom);
          });
        }
        
        return result;
      }
    } catch (error) {
      console.error('❌ Error during people picker query test:', error);
      return { success: false, error };
    }
  }

  /**
   * Enable people picker fields with proper SharePoint configuration
   * Call this method when you want to test people picker functionality
   */
  public enablePeoplePickerWithProperConfig(): void {
    this.enablePeoplePickerFields = true;
    console.log('✅ People picker fields enabled with proper SharePoint configuration');
    console.log('ℹ️ Note: This may require specific SharePoint list column configuration');
  }

  /**
   * Helper method to create query options with people picker field expansion
   * Enhanced to properly handle people picker fields based on feature flag
   */
  private createQueryOptions(additionalOptions?: Partial<IListQueryOptions>): IListQueryOptions {
    const baseOptions: IListQueryOptions = {};
    
    // Enable people picker field expansion when the feature is enabled
    if (this.enablePeoplePickerFields) {
      console.log('✅ Including people picker fields in query options');
      // For people picker fields ending with 'Id', expand without the 'Id' suffix
      baseOptions.expand = ['Responsibility', 'ByWhom'];
      
      // Add select fields to ensure we get the expanded user data
      // Use the base name (without Id) for expansion
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
      console.log('🚫 People picker fields disabled - excluding from query options');
      // Still select the fields but without expansion to get basic values
      baseOptions.select = ['*'];
    }
    
    return { ...baseOptions, ...additionalOptions };
  }

  /**
   * Create query options specifically for people picker testing
   * Use this method when you want to test people picker functionality
   */
  public createPeoplePickerTestQueryOptions(additionalOptions?: Partial<IListQueryOptions>): IListQueryOptions {
    console.log('🧪 Creating test query options for people picker fields');
    
    const testOptions: IListQueryOptions = {
      expand: ['Responsibility', 'ByWhom'],
      select: [
        'Id',
        'Title',
        'SelectType',
        'RiskDescription',
        'Responsibility/Id',
        'Responsibility/Title',
        'ResponsibilityId/Email',
        'ResponsibilityId/LoginName',
        'ByWhom/Id',
        'ByWhom/Title',
        'ByWhomId/Email',
        'ByWhomId/LoginName'
      ]
    };
    
    return { ...testOptions, ...additionalOptions };
  }

  /**
   * Convert IRaidItem to SharePoint format
   */
  private convertToSharePointItem(raidItem: IRaidItem): Omit<IRaidSharePointItem, 'Id'> {
    const spItem: any = {
      SelectType: raidItem.type, // Maps to SelectType internal field
      IdentificationDate: this.convertToISODate(raidItem.identificationDate),
      RiskDescription: raidItem.description || undefined, // Maps to RiskDescription internal field
      Applicability: raidItem.applicability || undefined,
      AssociatedGoal: raidItem.associatedGoal || undefined,
      RiskSource: raidItem.source || undefined, // Maps to RiskSource internal field
      RiskCategory: raidItem.category || undefined, // Maps to RiskCategory internal field
      Impact: raidItem.impact || undefined,
      RiskPriority: raidItem.priority || undefined, // Maps to RiskPriority internal field
      
      // Risk specific - handle numeric fields properly
      ImpactValue: this.convertToNumber(raidItem.impactValue),
      ProbabilityValue: this.convertToNumber(raidItem.probabilityValue),
      RiskExposure: this.convertToNumber(raidItem.riskExposure),
      
      // Opportunity specific
      PotentialCost: this.convertToNumber(raidItem.potentialCost),
      PotentialBenefit: this.convertToNumber(raidItem.potentialBenefit),
      OpportunityValue: this.convertToNumber(raidItem.opportunityValue),
      TypeOfAction: raidItem.typeOfAction || undefined,
      ActionPlan: raidItem.actionPlan || undefined,
      TargetDate: this.convertToISODate(raidItem.targetDate),
      ActualDate: this.convertToISODate(raidItem.actualDate),
      RiskStatus: raidItem.status || undefined, // Maps to RiskStatus internal field
      
      // Issue/Assumption/Dependency specific
      IssueDetails: raidItem.details || undefined, // Maps to IssueDetails internal field
      IDADate: this.convertToISODate(raidItem.date), // Maps to IDADate internal field
      ImplementationActions: raidItem.implementationActions || undefined,
      PlannedClosureDate: this.convertToISODate(raidItem.plannedClosureDate),
      ActualClosureDate: this.convertToISODate(raidItem.actualClosureDate),
      
      // Common
      Effectiveness: raidItem.effectiveness || undefined,
      Remarks: raidItem.remarks || undefined
    };

    debugger;
    // Handle people picker fields separately with feature flag and enhanced error handling
    if (this.enablePeoplePickerFields) {
      console.log('Processing people picker fields...');
      
      try {
        if (raidItem.responsibility) {
          console.log('Raw responsibility value:', raidItem.responsibility);
          const responsibilityValue = this.convertUserFieldForSharePoint(raidItem.responsibility);
          if (responsibilityValue !== null) {
            spItem.ResponsibilityId = responsibilityValue;
            console.log('✅ Successfully added ResponsibilityId field:', responsibilityValue);
          } else {
            console.log('⚠️ ResponsibilityId field converted to null, skipping');
          }
        } else {
          console.log('No responsibility value provided');
        }
      } catch (error) {
        console.error('❌ Error processing ResponsibilityId field:', error);
        console.warn('Disabling people picker fields due to ResponsibilityId field error');
        this.enablePeoplePickerFields = false;
      }

      try {
        if (raidItem.byWhom) {
          console.log('Raw byWhom value:', raidItem.byWhom);
          const byWhomValue = this.convertUserFieldForSharePoint(raidItem.byWhom);
          if (byWhomValue !== null) {
            spItem.ByWhomId = byWhomValue;
            console.log('✅ Successfully added ByWhomId field:', byWhomValue);
          } else {
            console.log('⚠️ ByWhomId field converted to null, skipping');
          }
        } else {
          console.log('No byWhom value provided');
        }
      } catch (error) {
        console.error('❌ Error processing ByWhomId field:', error);
        console.warn('Disabling people picker fields due to ByWhomId field error');
        this.enablePeoplePickerFields = false;
      }
    } else {
      console.log('🚫 People picker fields disabled, skipping Responsibility and ByWhom fields');
    }

    // Remove undefined fields to avoid SharePoint issues
    Object.keys(spItem).forEach(key => {
      if (spItem[key as keyof typeof spItem] === undefined) {
        delete spItem[key as keyof typeof spItem];
      }
    });

    // Log each field and its type for debugging
    console.log('SharePoint item field analysis:');
    Object.keys(spItem).forEach(key => {
      const value = spItem[key];
      console.log(`  ${key}: ${JSON.stringify(value)} (type: ${typeof value})`);
    });

    return spItem;
  }

  /**
   * Convert SharePoint item to IRaidItem format
   */
  private convertFromSharePointItem(spItem: IRaidSharePointItem): IExtendedRaidItem {
    debugger;
    const raidItem: IExtendedRaidItem = {
      id: spItem.Id || 0,
      type: spItem.SelectType, // Maps from SelectType internal field
      identificationDate: spItem.IdentificationDate,
      description: spItem.RiskDescription, // Maps from RiskDescription internal field
      applicability: spItem.Applicability,
      associatedGoal: spItem.AssociatedGoal,
      source: spItem.RiskSource, // Maps from RiskSource internal field
      category: spItem.RiskCategory, // Maps from RiskCategory internal field
      impact: spItem.Impact,
      priority: spItem.RiskPriority, // Maps from RiskPriority internal field
      
      // Risk specific
      impactValue: spItem.ImpactValue,
      probabilityValue: spItem.ProbabilityValue,
      riskExposure: spItem.RiskExposure,
      
      // Opportunity specific
      potentialCost: spItem.PotentialCost,
      potentialBenefit: spItem.PotentialBenefit,
      opportunityValue: spItem.OpportunityValue,
      typeOfAction: spItem.TypeOfAction,
      actionPlan: spItem.ActionPlan,
      // Check expanded field first (Responsibility), then fall back to ResponsibilityId
      responsibility: this.parsePeoplePickerField((spItem as any).Responsibility || spItem.ResponsibilityId?.[0]),
      targetDate: spItem.TargetDate,
      actualDate: spItem.ActualDate,
      status: spItem.RiskStatus, // Maps from RiskStatus internal field
      
      // Issue/Assumption/Dependency specific
      details: spItem.IssueDetails, // Maps from IssueDetails internal field
      date: spItem.IDADate, // Maps from IDADate internal field
      // Check expanded field first (ByWhom), then fall back to ByWhomId
      byWhom: this.parsePeoplePickerField((spItem as any).ByWhom || spItem.ByWhomId?.[0]),
      implementationActions: spItem.ImplementationActions,
      plannedClosureDate: spItem.PlannedClosureDate,
      actualClosureDate: spItem.ActualClosureDate,
      
      // Common
      effectiveness: spItem.Effectiveness,
      remarks: spItem.Remarks
    };

    return raidItem;
  }

  /**
   * Helper method to parse SharePoint people picker field responses
   * Enhanced to handle various SharePoint response formats
   * Returns array of IPersonPickerUser for multi-value fields, or single user wrapped in array
   */
  private parsePeoplePickerField(value: any): any {
    if (!value) return undefined;
    
    try {
      console.log('Parsing people picker field from SharePoint:', value);
      
      // Handle different SharePoint response formats
      if (typeof value === 'object') {
        // Handle array of users (multi-valued people picker)
        if (Array.isArray(value)) {
          const parsedUsers = value.map((user: any) => this.parseUserObject(user)).filter((user: any) => user !== null && user !== undefined);
          console.log('Parsed array of users:', parsedUsers);
          return parsedUsers.length > 0 ? parsedUsers : undefined;
        }
        
        // Handle single user object or expanded user data
        if (value.results && Array.isArray(value.results)) {
          // SharePoint OData response format with results array
          const parsedUsers = value.results.map((user: any) => this.parseUserObject(user)).filter((user: any) => user !== null && user !== undefined);
          console.log('Parsed users from results array:', parsedUsers);
          return parsedUsers.length > 0 ? parsedUsers : undefined;
        }
        
        // Handle single user object - wrap in array for consistency
        const parsedUser = this.parseUserObject(value);
        if (parsedUser) {
          console.log('Parsed single user, wrapping in array:', [parsedUser]);
          return [parsedUser];
        }
      }
      
      // If it's a string, try to parse as JSON first
      if (typeof value === 'string') {
        try {
          const parsed = JSON.parse(value);
          return this.parsePeoplePickerField(parsed); // Recursive call with parsed object
        } catch {
          // If JSON parsing fails, treat as simple string
          // Could be email, login name, or user ID as string
          if (value.indexOf('@') !== -1) {
            // Looks like an email - wrap in array
            return [{
              id: value,
              email: value,
              loginName: value,
              displayName: value.split('@')[0]
            }];
          } else {
            // Could be user ID or other identifier
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
      
      // If it's a number, treat as user ID - wrap in array
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
      console.warn('Error parsing people picker field:', error, value);
      return undefined;
    }
  }

  /**
   * Helper method to parse individual user objects from SharePoint responses
   * Returns user object in IPersonPickerUser format: { id, loginName, displayName, email }
   */
  private parseUserObject(user: any): any {
    if (!user || typeof user !== 'object') return user;
    
    try {
      // Create a standardized user object matching IPersonPickerUser interface
      const parsedUser: any = {};
      
      // Handle user ID - convert to string as per IPersonPickerUser interface
      if (user.Id !== undefined) parsedUser.id = String(user.Id);
      else if (user.id !== undefined) parsedUser.id = String(user.id);
      else if (user.ID !== undefined) parsedUser.id = String(user.ID);
      
      // Handle email
      if (user.Email) parsedUser.email = user.Email;
      else if (user.email) parsedUser.email = user.email;
      else if (user.EMail) parsedUser.email = user.EMail;
      
      // Handle login name
      if (user.LoginName) parsedUser.loginName = user.LoginName;
      else if (user.loginName) parsedUser.loginName = user.loginName;
      else if (user.UserName) parsedUser.loginName = user.UserName;
      else if (user.userName) parsedUser.loginName = user.userName;
      
      // Handle display name (map Title to displayName for IPersonPickerUser interface)
      if (user.Title) parsedUser.displayName = user.Title;
      else if (user.DisplayName) parsedUser.displayName = user.DisplayName;
      else if (user.displayName) parsedUser.displayName = user.displayName;
      else if (user.text) parsedUser.displayName = user.text;
      else if (user.name) parsedUser.displayName = user.name;
      else if (user.Name) parsedUser.displayName = user.Name;
      else if (user.title) parsedUser.displayName = user.title;
      
      // Fallback: if no id, try to extract from other fields
      if (!parsedUser.id) {
        if (parsedUser.loginName) parsedUser.id = parsedUser.loginName;
        else if (parsedUser.email) parsedUser.id = parsedUser.email;
      }
      
      // Ensure all required fields are present for IPersonPickerUser
      if (!parsedUser.loginName && parsedUser.email) {
        parsedUser.loginName = parsedUser.email;
      }
      if (!parsedUser.email && parsedUser.loginName && parsedUser.loginName.indexOf('@') !== -1) {
        parsedUser.email = parsedUser.loginName;
      }
      if (!parsedUser.displayName && parsedUser.email) {
        parsedUser.displayName = parsedUser.email.split('@')[0];
      }
      
      // If we couldn't extract any meaningful data, return the original
      if (!parsedUser.id && !parsedUser.email && !parsedUser.loginName) {
        console.warn('Could not parse user object, returning original:', user);
        return user;
      }
      
      console.log('Successfully parsed user object to IPersonPickerUser format:', parsedUser);
      return parsedUser;
    } catch (error) {
      console.warn('Error parsing user object:', error, user);
      return user;
    }
  }

  /**
   * Helper method to convert user fields for SharePoint people picker
   * Based on the attached solution pattern for handling UsersId format
   * Prioritizes user IDs for both byWhom and responsibility fields
   */
  private convertUserFieldForSharePoint(userField: any): any {
    if (!userField) return null;
    
    try {
      console.log('Converting user field for SharePoint:', userField);
      
      // SharePoint people picker fields can be set using different formats:
      // 1. User ID number (most reliable)
      // 2. Email address
      // 3. Login name
      // 4. UsersId format for specific user ID setting (as per attached solution)
      
      // If it's already a number (user ID), use it directly - most reliable method
      if (typeof userField === 'number') {
        console.log('Using user ID directly:', userField);
        return userField;
      }
      
      // If it's a string, determine the format
      if (typeof userField === 'string') {
        // Check if it's an email address
        if (userField.indexOf('@') !== -1) {
          console.log('Using email address:', userField);
          return userField;
        }
        
        // If it has the membership prefix, extract the email
        if (userField.indexOf('i:0#.f|membership|') !== -1) {
          const email = userField.replace('i:0#.f|membership|', '');
          if (email.indexOf('@') !== -1) {
            console.log('Extracted email from login name:', email);
            return email;
          }
        }
        
        // If it's a numeric string, convert to number (user ID)
        const numericValue = parseInt(userField, 10);
        if (!isNaN(numericValue) && numericValue > 0) {
          console.log('Converting numeric string to user ID:', numericValue);
          return numericValue;
        }
        
        console.log('Using string value as-is:', userField);
        return userField;
      }
      
      // If it's an array (multi-valued people picker)
      if (Array.isArray(userField)) {
        const processedUsers = [];
        
        for (const user of userField) {
          let processedUser = null;
          
          if (typeof user === 'number') {
            processedUser = user;
          } else if (typeof user === 'string') {
            if (user.indexOf('@') !== -1) {
              processedUser = user;
            } else if (user.indexOf('i:0#.f|membership|') !== -1) {
              const email = user.replace('i:0#.f|membership|', '');
              processedUser = email.indexOf('@') !== -1 ? email : null;
            } else {
              const numericValue = parseInt(user, 10);
              if (!isNaN(numericValue) && numericValue > 0) {
                processedUser = numericValue;
              } else {
                processedUser = user;
              }
            }
          } else if (typeof user === 'object' && user !== null) {
            // Handle people picker object format
            processedUser = this.extractUserIdFromObject(user);
          }
          
          if (processedUser !== null) {
            processedUsers.push(processedUser);
          }
        }
        
        console.log('Processed user array:', processedUsers);
        return processedUsers.length > 0 ? processedUsers : null;
      }
      
      // If it's a single user object (people picker result)
      if (typeof userField === 'object' && userField !== null) {
        const extractedUser = this.extractUserIdFromObject(userField);
        console.log('Extracted user from object:', extractedUser);
        return extractedUser;
      }
      
      console.warn('Unable to convert user field to SharePoint format:', userField);
      return null;
    } catch (error) {
      console.warn('Error converting user field for SharePoint:', error, userField);
      return null;
    }
  }

  /**
   * Helper method to extract user ID or email from people picker object
   * Prioritizes user IDs for both byWhom and responsibility fields
   */
  private extractUserIdFromObject(userObj: any): any {
    if (!userObj || typeof userObj !== 'object') return null;
    
    try {
      // Priority order for user identification:
      // 1. User ID (most reliable - always preferred for byWhom and responsibility)
      // 2. Numeric values that can be converted to user ID
      // 3. Email address (fallback only)
      // 4. Login name with email extraction (fallback only)
      
      // Check for user ID in various property names - HIGHEST PRIORITY
      if (userObj.id && typeof userObj.id === 'number' && userObj.id > 0) {
        console.log('Found user ID in "id" property:', userObj.id);
        return userObj.id;
      }
      if (userObj.userId && typeof userObj.userId === 'number' && userObj.userId > 0) {
        console.log('Found user ID in "userId" property:', userObj.userId);
        return userObj.userId;
      }
      if (userObj.Id && typeof userObj.Id === 'number' && userObj.Id > 0) {
        console.log('Found user ID in "Id" property:', userObj.Id);
        return userObj.Id;
      }
      if (userObj.ID && typeof userObj.ID === 'number' && userObj.ID > 0) {
        console.log('Found user ID in "ID" property:', userObj.ID);
        return userObj.ID;
      }
      
      // Check for numeric string values that can be converted to user ID
      if (userObj.id && typeof userObj.id === 'string') {
        const numericId = parseInt(userObj.id, 10);
        if (!isNaN(numericId) && numericId > 0) {
          console.log('Converted string id to user ID:', numericId);
          return numericId;
        }
      }
      if (userObj.userId && typeof userObj.userId === 'string') {
        const numericId = parseInt(userObj.userId, 10);
        if (!isNaN(numericId) && numericId > 0) {
          console.log('Converted string userId to user ID:', numericId);
          return numericId;
        }
      }
      if (userObj.Id && typeof userObj.Id === 'string') {
        const numericId = parseInt(userObj.Id, 10);
        if (!isNaN(numericId) && numericId > 0) {
          console.log('Converted string Id to user ID:', numericId);
          return numericId;
        }
      }
      
      // Check for email address
      if (userObj.email && userObj.email.indexOf('@') !== -1) {
        return userObj.email;
      }
      if (userObj.Email && userObj.Email.indexOf('@') !== -1) {
        return userObj.Email;
      }
      
      // Check for login name and extract email
      if (userObj.loginName) {
        if (userObj.loginName.indexOf('i:0#.f|membership|') !== -1) {
          const email = userObj.loginName.replace('i:0#.f|membership|', '');
          if (email.indexOf('@') !== -1) {
            return email;
          }
        } else if (userObj.loginName.indexOf('@') !== -1) {
          return userObj.loginName;
        }
      }
      if (userObj.LoginName) {
        if (userObj.LoginName.indexOf('i:0#.f|membership|') !== -1) {
          const email = userObj.LoginName.replace('i:0#.f|membership|', '');
          if (email.indexOf('@') !== -1) {
            return email;
          }
        } else if (userObj.LoginName.indexOf('@') !== -1) {
          return userObj.LoginName;
        }
      }
      
      // Check for other text representations
      if (userObj.text && userObj.text.indexOf('@') !== -1) {
        return userObj.text;
      }
      if (userObj.title && userObj.title.indexOf('@') !== -1) {
        return userObj.title;
      }
      if (userObj.Title && userObj.Title.indexOf('@') !== -1) {
        return userObj.Title;
      }
      
      // Check for key property (often used in people picker controls)
      // Prioritize numeric keys (user IDs) over email keys
      if (userObj.key) {
        // First check if key is a numeric user ID
        const numericValue = parseInt(userObj.key, 10);
        if (!isNaN(numericValue) && numericValue > 0) {
          console.log('Found numeric user ID in key property:', numericValue);
          return numericValue;
        }
        
        // Then check for login name format
        if (userObj.key.indexOf('i:0#.f|membership|') !== -1) {
          const email = userObj.key.replace('i:0#.f|membership|', '');
          if (email.indexOf('@') !== -1) {
            console.log('Extracted email from key login name:', email);
            return email;
          }
        } else if (userObj.key.indexOf('@') !== -1) {
          console.log('Found email in key property:', userObj.key);
          return userObj.key;
        }
      }
      
      console.warn('Could not extract user identifier from object:', userObj);
      return null;
    } catch (error) {
      console.warn('Error extracting user from object:', error, userObj);
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
      
      // If it's a number (timestamp), convert it
      if (typeof value === 'number' && !isNaN(value)) {
        const date = new Date(value);
        if (!isNaN(date.getTime())) {
          return date.toISOString();
        }
      }
      
      return undefined;
    } catch (error) {
      console.warn('Error converting date value:', value, error);
      return undefined;
    }
  }

  /**
   * CREATE: Add new RAID item
   */
  async createRaidItem(raidItem: Omit<IRaidItem, 'id'>): Promise<IExtendedRaidItem | null> {
    try {

      debugger;
      console.log('Creating RAID item with original data:', JSON.stringify(raidItem, null, 2));
      
      const spItem = this.convertToSharePointItem({ ...raidItem, id: 0 });
      console.log('Converted SharePoint item data:', JSON.stringify(spItem, null, 2));
      
      const result = await this.spService.createItem<IRaidSharePointItem>(
        { listName: this.listName },
        spItem
      );

      if (result.success && result.data) {
        console.log('RAID item created successfully:', result.data);
        return this.convertFromSharePointItem(result.data);
      } else {
        console.error(ERROR_MESSAGES.CREATE_FAILED, {
          error: result.error,
          errorDetails: result.errorDetails,
          originalItem: raidItem,
          convertedItem: spItem
        });
        return null;
      }
    } catch (error) {
      console.error('Error creating RAID item:', {
        error,
        originalItem: raidItem
      });
      return null;
    }
  }

  /**
   * READ: Get all RAID items
   */
  async getAllRaidItems(): Promise<IExtendedRaidItem[]> {
    try {
      const options = this.createQueryOptions({
        orderBy: 'Modified'
      });

      const result = await this.spService.getItems<IRaidSharePointItem>(
        { listName: this.listName },
        options
      );

      if (result.success && result.data) {
        return result.data.map(spItem => this.convertFromSharePointItem(spItem));
      } else {
        console.error('Failed to get RAID items:', result.error);
        return [];
      }
    } catch (error) {
      console.error('Error getting RAID items:', error);
      return [];
    }
  }

  /**
   * READ: Get RAID items by type
   */
  async getRaidItemsByType(type: RaidType): Promise<IExtendedRaidItem[]> {
    try {
      const options = this.createQueryOptions({
        filter: `SelectType eq '${type}'`, // Updated to use SelectType internal field
        orderBy: 'Modified'
      });

      const result = await this.spService.getItems<IRaidSharePointItem>(
        { listName: this.listName },
        options
      );

      if (result.success && result.data) {
        return result.data.map(spItem => this.convertFromSharePointItem(spItem));
      } else {
        console.error(`Failed to get ${type} items:`, result.error);
        return [];
      }
    } catch (error) {
      console.error(`Error getting ${type} items:`, error);
      return [];
    }
  }

  /**
   * READ: Get RAID item by ID
   */
  async getRaidItemById(itemId: number): Promise<IExtendedRaidItem | null> {
    try {
      const options = this.createQueryOptions();
      
      const result = await this.spService.getItemById<IRaidSharePointItem>(
        { listName: this.listName },
        itemId,
        options
      );

      if (result.success && result.data) {
        return this.convertFromSharePointItem(result.data);
      } else {
        console.error(`Failed to get RAID item ${itemId}:`, result.error);
        return null;
      }
    } catch (error) {
      console.error(`Error getting RAID item ${itemId}:`, error);
      return null;
    }
  }

  /**
   * UPDATE: Update RAID item
   */
  async updateRaidItem(itemId: number, updates: Partial<IRaidItem>): Promise<IExtendedRaidItem | null> {
    try {
      // Get current item first to merge with updates
      const currentItem = await this.getRaidItemById(itemId);
      if (!currentItem) {
        console.error('Item not found for update');
        return null;
      }

      // Merge updates with current item
      const updatedItem = { ...currentItem, ...updates };
      const spUpdates = this.convertToSharePointItem(updatedItem);

      const result = await this.spService.updateItem<IRaidSharePointItem>(
        { listName: this.listName },
        itemId,
        spUpdates
      );

      if (result.success && result.data) {
        return this.convertFromSharePointItem(result.data);
      } else {
        console.error(`Failed to update RAID item ${itemId}:`, result.error);
        return null;
      }
    } catch (error) {
      console.error(`Error updating RAID item ${itemId}:`, error);
      return null;
    }
  }

  /**
   * DELETE: Delete RAID item
   */
  async deleteRaidItem(itemId: number): Promise<boolean> {
    try {
      const result = await this.spService.deleteItem(
        { listName: this.listName },
        itemId
      );

      if (result.success) {
        console.log(`RAID item ${itemId} deleted successfully`);
        return true;
      } else {
        console.error(`Failed to delete RAID item ${itemId}:`, result.error);
        return false;
      }
    } catch (error) {
      console.error(`Error deleting RAID item ${itemId}:`, error);
      return false;
    }
  }

  /**
   * UTILITY: Get RAID items by status
   */
  async getRaidItemsByStatus(status: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `RiskStatus eq '${status}'`, // Updated to use RiskStatus internal field
        orderBy: 'Modified'
      };

      const result = await this.spService.getItems<IRaidSharePointItem>(
        { listName: this.listName },
        options
      );

      if (result.success && result.data) {
        return result.data.map(spItem => this.convertFromSharePointItem(spItem));
      } else {
        console.error(`Failed to get items with status ${status}:`, result.error);
        return [];
      }
    } catch (error) {
      console.error(`Error getting items with status ${status}:`, error);
      return [];
    }
  }

  /**
   * UTILITY: Get RAID items by priority
   */
  async getRaidItemsByPriority(priority: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `RiskPriority eq '${priority}'`, // Updated to use RiskPriority internal field
        orderBy: 'Modified'
      };

      const result = await this.spService.getItems<IRaidSharePointItem>(
        { listName: this.listName },
        options
      );

      if (result.success && result.data) {
        return result.data.map(spItem => this.convertFromSharePointItem(spItem));
      } else {
        console.error(`Failed to get items with priority ${priority}:`, result.error);
        return [];
      }
    } catch (error) {
      console.error(`Error getting items with priority ${priority}:`, error);
      return [];
    }
  }

  /**
   * UTILITY: Search RAID items
   */
  async searchRaidItems(searchTerm: string): Promise<IExtendedRaidItem[]> {
    try {
      const options: IListQueryOptions = {
        filter: `substringof('${searchTerm}', Title) or substringof('${searchTerm}', RiskDescription) or substringof('${searchTerm}', IssueDetails)`, // Updated to use internal field names
        orderBy: 'Modified'
      };

      const result = await this.spService.getItems<IRaidSharePointItem>(
        { listName: this.listName },
        options
      );

      if (result.success && result.data) {
        return result.data.map(spItem => this.convertFromSharePointItem(spItem));
      } else {
        console.error(`Failed to search items with term ${searchTerm}:`, result.error);
        return [];
      }
    } catch (error) {
      console.error(`Error searching items with term ${searchTerm}:`, error);
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

  /**
   * UTILITY: Bulk update status
   */
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
      console.error('Error in bulk status update:', error);
      return false;
    }
  }

  /**
   * UTILITY: Check if RAID list exists
   */
  async checkListExists(): Promise<boolean> {
    return await this.spService.listExists({ listName: this.listName });
  }

  /**
   * UTILITY: Get list information
   */
  async getListInfo(): Promise<any> {
    const result = await this.spService.getListInfo({ listName: this.listName });
    return result.success ? result.data : null;
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