import * as React from 'react';
import { 
  Modal, 
  TextField, 
  DatePicker, 
  Dropdown, 
  PrimaryButton, 
  DefaultButton,
  IconButton,
  Text,
  Checkbox,
  Pivot,
  PivotItem
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './RaidForm.module.scss';
import { IRaidItem, RaidType, IRaidAction, IPersonPickerUser } from '../IRaidItem';

export interface IRaidFormProps {
  isOpen: boolean;
  type: RaidType;
  item: IRaidItem | null;
  onSave: (item: IRaidItem) => Promise<void>;
  onCancel: () => void;
  context: WebPartContext;
}

const RaidForm: React.FC<IRaidFormProps> = ({ isOpen, type, item, onSave, onCancel, context }) => {
  const [formData, setFormData] = React.useState<Partial<IRaidItem>>(item ? { ...item } : { type });
  const [selectedActionTypes, setSelectedActionTypes] = React.useState<string[]>([]);
  const [mitigationAction, setMitigationAction] = React.useState<IRaidAction | null>(null);
  const [contingencyAction, setContingencyAction] = React.useState<IRaidAction | null>(null);

  // Helper function to convert PeoplePicker items to our format
  // PeoplePicker onChange returns items in format: { id: string, loginName: string, text: string, secondaryText: string }
  debugger;
  const convertPeoplePickerItems = (items: any[]): IPersonPickerUser[] => {
    return items.map(item => ({
      id: item.id || item.loginName,
      loginName: item.loginName,
      displayName: item.text || item.displayName,
      email: item.secondaryText || item.mail || item.email
    }));
  };

  // Helper function to convert SharePoint user data back to PeoplePicker format for display
  // SharePoint returns users with { id, loginName, displayName, email }
  // PeoplePicker's defaultSelectedUsers expects user IDs (as strings) when email/loginName not available
  const convertToPickerFormat = (users: IPersonPickerUser[] | any): string[] => {
    console.log('🔄 Converting to picker format, input:', users);
    
    if (!users) {
      console.log('🔄 No users to convert');
      return [];
    }
    
    // Ensure users is an array
    const userArray = Array.isArray(users) ? users : [users];
    console.log('🔄 User array:', userArray);
    
    const result = userArray.map(user => {
      console.log('🔄 Processing user:', user);
      
      // Priority 1: Email address (most reliable for PeoplePicker)
      if (user.email && user.email.indexOf('@') !== -1) {
        console.log('🔄 Using email:', user.email);
        return user.email;
      }
      
      // Priority 2: LoginName with @ symbol
      if (user.loginName && user.loginName.indexOf('@') !== -1) {
        console.log('🔄 Using loginName:', user.loginName);
        return user.loginName;
      }
      
      // Priority 3: Extract email from SharePoint membership format loginName
      if (user.loginName && user.loginName.indexOf('i:0#.f|membership|') !== -1) {
        const extracted = user.loginName.replace('i:0#.f|membership|', '');
        console.log('🔄 Extracted email from loginName:', extracted);
        return extracted;
      }
      
      // Priority 4: Any loginName
      if (user.loginName) {
        console.log('🔄 Using loginName fallback:', user.loginName);
        return user.loginName;
      }
      
      // Priority 5: Use displayName as fallback (PeoplePicker can resolve by name)
      if (user.displayName) {
        console.log('🔄 Using displayName:', user.displayName);
        return user.displayName;
      }
      
      // Priority 6: Use id as last resort (will be resolved by PeoplePicker)
      if (user.id) {
        console.log('🔄 Using id fallback:', user.id);
        return user.id;
      }
      
      console.log('🔄 No valid identifier found');
      return null;
    }).filter(item => item !== null) as string[];
    
    console.log('🔄 Final converted result:', result);
    return result;
  };

  // Helper function to convert IPersonPickerUser array to user IDs for SharePoint
  // Used when saving to SharePoint (Id suffix fields require numeric IDs)
  const convertToUserIds = (users: IPersonPickerUser[]): number[] => {
    if (!users || !Array.isArray(users)) return [];
    
    return users.map(user => {
      // Extract numeric user ID from string if needed
      const userId = typeof user.id === 'string' ? parseInt(user.id, 10) : user.id;
      return typeof userId === 'number' && !isNaN(userId) ? userId : null;
    }).filter(id => id !== null) as number[];
  };

  React.useEffect(() => {
    setFormData(item ? { ...item } : { type });
    
    // Debug: Log the item data to see what we're receiving
    if (item) {
      console.log('📝 RaidForm - Item data received:', item);
      console.log('📝 RaidForm - ByWhom data:', item.byWhom);
      console.log('📝 RaidForm - ByWhom converted for picker:', item.byWhom ? convertToPickerFormat(item.byWhom) : []);
      console.log('📝 RaidForm - Responsibility data:', item.responsibility);
      console.log('📝 RaidForm - Responsibility converted for picker:', item.responsibility ? convertToPickerFormat(item.responsibility) : []);
    }
    
    // Initialize action type selection and separate actions
    if (item?.actions && item.actions.length > 0) {
      let mitigation: IRaidAction | undefined;
      let contingency: IRaidAction | undefined;
      
      for (let i = 0; i < item.actions.length; i++) {
        if (item.actions[i].type === 'Mitigation') {
          mitigation = item.actions[i];
        } else if (item.actions[i].type === 'Contingency') {
          contingency = item.actions[i];
        }
      }
      
      setMitigationAction(mitigation || null);
      setContingencyAction(contingency || null);
      
      const types: string[] = [];
      if (mitigation) types.push('Mitigation');
      if (contingency) types.push('Contingency');
      setSelectedActionTypes(types);
    } else {
      setMitigationAction(null);
      setContingencyAction(null);
      setSelectedActionTypes([]);
    }
  }, [item, type]);

  // Calculate OpportunityValue when potentialCost or potentialBenefit changes
  React.useEffect(() => {
    if (type === 'Opportunity' && (formData.potentialCost || formData.potentialBenefit)) {
      const cost = Number(formData.potentialCost) || 0;
      const benefit = Number(formData.potentialBenefit) || 0;
      const calculatedValue = cost * benefit;
      if (formData.opportunityValue !== calculatedValue) {
        updateFormData('opportunityValue', calculatedValue);
      }
    }
  }, [formData.potentialCost, formData.potentialBenefit, type]);

  // Calculate RiskExposure when impactValue or probabilityValue changes
  React.useEffect(() => {
    if (type === 'Risk' && (formData.impactValue || formData.probabilityValue)) {
      const impact = Number(formData.impactValue) || 0;
      const probability = Number(formData.probabilityValue) || 0;
      const calculatedValue = impact * probability;
      if (formData.riskExposure !== calculatedValue) {
        updateFormData('riskExposure', calculatedValue);
      }
    }
  }, [formData.impactValue, formData.probabilityValue, type]);

  const updateFormData = (field: string, value: any): void => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };



  const handleActionTypeChange = (actionType: string, checked: boolean): void => {
    if (checked) {
      setSelectedActionTypes(prev => [...prev, actionType]);
      // Initialize with one empty action if none exist
      if (actionType === 'Mitigation' && !mitigationAction) {
        setMitigationAction(createEmptyAction('Mitigation'));
      } else if (actionType === 'Contingency' && !contingencyAction) {
        setContingencyAction(createEmptyAction('Contingency'));
      }
    } else {
      setSelectedActionTypes(prev => prev.filter(type => type !== actionType));
      // Clear actions for this type
      if (actionType === 'Mitigation') {
        setMitigationAction(null);
      } else if (actionType === 'Contingency') {
        setContingencyAction(null);
      }
    }
  };

  const createEmptyAction = (type: string): IRaidAction => ({
    type,
    plan: '',
    responsibility: [],
    targetDate: '',
    actualDate: '',
    status: ''
  });





  const updateActionInType = (actionType: string, field: string, value: string | IPersonPickerUser[]): void => {
    if (actionType === 'Mitigation' && mitigationAction) {
      setMitigationAction({ ...mitigationAction, [field]: value });
    } else if (actionType === 'Contingency' && contingencyAction) {
      setContingencyAction({ ...contingencyAction, [field]: value });
    }
  };



  const handleSave = async (): Promise<void> => {
    const allActions: IRaidAction[] = [];
    if (type === 'Risk') {
      // Convert responsibility arrays to user IDs for Risk actions
      if (mitigationAction) {
        const mitigationToSave = {
          ...mitigationAction,
          responsibility: mitigationAction.responsibility 
            ? convertToUserIds(mitigationAction.responsibility) as any
            : []
        };
        allActions.push(mitigationToSave);
      }
      if (contingencyAction) {
        const contingencyToSave = {
          ...contingencyAction,
          responsibility: contingencyAction.responsibility 
            ? convertToUserIds(contingencyAction.responsibility) as any
            : []
        };
        allActions.push(contingencyToSave);
      }
    }
    
    // Prepare item to save with user IDs for people picker fields
    const itemToSave: any = {
      ...formData,
      type,
      id: formData.id || 0,
      actions: type === 'Risk' ? allActions : undefined
    };

    // Convert responsibility and byWhom fields to user IDs for SharePoint
    if (formData.responsibility) {
      itemToSave.responsibility = convertToUserIds(formData.responsibility);
    }
    if (formData.byWhom) {
      itemToSave.byWhom = convertToUserIds(formData.byWhom);
    }
    
    await onSave(itemToSave);
  };

  const renderFormFields = (): React.ReactElement => {
    if (type === 'Issue' || type === 'Assumption' || type === 'Dependency') {
      return renderIssueForm();
    } else if (type === 'Opportunity') {
      return renderOpportunityForm();
    } else if (type === 'Risk') {
      return renderRiskForm();
    }
    
    return <div>Form not implemented for type: {type}</div>;
  };

  const renderIssueForm = (): React.ReactElement => {
    
    return (
      <div>
        <TextField
          label="Details"
          multiline
          rows={3}
          value={formData.details || ''}
          onChange={(_, value) => updateFormData('details', value || '')}
          placeholder="Enter value here"
        />
        
        <DatePicker
          label="Date"
          value={formData.date ? new Date(formData.date) : undefined}
          onSelectDate={(date) => updateFormData('date', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <PeoplePicker
          key={`byWhom-${formData.id || 'new'}-${formData.byWhom ? JSON.stringify(formData.byWhom) : 'empty'}`}
          context={{
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient
          }}
          titleText="By Whom (Name)"
          personSelectionLimit={3}
          groupName=""
          showtooltip={true}
          required={false}
          disabled={false}
          onChange={(items: any[]) => {
            console.log('👤 ByWhom PeoplePicker onChange:', items);
            const selectedUsers = convertPeoplePickerItems(items);
            console.log('👤 ByWhom converted users:', selectedUsers);
            updateFormData('byWhom', selectedUsers);
          }}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
          defaultSelectedUsers={formData.byWhom ? convertToPickerFormat(formData.byWhom) : []}
        />
        
        <TextField
          label="Implementation Actions"
          multiline
          rows={3}
          value={formData.implementationActions || ''}
          onChange={(_, value) => updateFormData('implementationActions', value || '')}
          placeholder="Enter value here"
        />
        
        <PeoplePicker
          key={`responsibility-issue-${formData.id || 'new'}-${formData.responsibility ? JSON.stringify(formData.responsibility) : 'empty'}`}
          context={{
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient
          }}
          titleText="Responsibility"
          personSelectionLimit={3}
          groupName=""
          showtooltip={true}
          required={false}
          disabled={false}
          onChange={(items: any[]) => {
            console.log('👤 Responsibility (Issue) PeoplePicker onChange:', items);
            const selectedUsers = convertPeoplePickerItems(items);
            console.log('👤 Responsibility (Issue) converted users:', selectedUsers);
            updateFormData('responsibility', selectedUsers);
          }}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
          defaultSelectedUsers={formData.responsibility ? convertToPickerFormat(formData.responsibility) : []}
        />
        
        <DatePicker
          label="Planned Closure Date"
          value={formData.plannedClosureDate ? new Date(formData.plannedClosureDate) : undefined}
          onSelectDate={(date) => updateFormData('plannedClosureDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <DatePicker
          label="Actual Closure Date"
          value={formData.actualClosureDate ? new Date(formData.actualClosureDate) : undefined}
          onSelectDate={(date) => updateFormData('actualClosureDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <TextField
          label="Remarks"
          multiline
          rows={3}
          value={formData.remarks || ''}
          onChange={(_, value) => updateFormData('remarks', value || '')}
          placeholder="Enter value here"
        />
      </div>
    );
  }

  const renderOpportunityForm = (): React.ReactElement => {
    
    const dropdownOptions = {
      associatedGoal: [
        { key: 'BG01', text: 'BG01' },
        { key: 'BG02', text: 'BG02' }
      ],
      source: [
        { key: 'Internal', text: 'Internal' },
        { key: 'External', text: 'External' }
      ],
      category: [
        { key: 'Resource', text: 'Resource' },
        { key: 'Customer', text: 'Customer' },
        { key: 'Information Security', text: 'Information Security' },
        { key: 'Technology', text: 'Technology' },
        { key: 'Business', text: 'Business' },
        { key: 'Process', text: 'Process' },
        { key: 'Others', text: 'Others' }
      ],
      priority: [
        { key: 'High', text: 'High' },
        { key: 'Medium', text: 'Medium' },
        { key: 'Low', text: 'Low' }
      ],
      status: [
        { key: 'Open', text: 'Open' },
        { key: 'Closed', text: 'Closed' },
        { key: 'In Progress', text: 'In Progress' },
        { key: 'Monitoring', text: 'Monitoring' }
      ],
      typeOfAction: [
        { key: 'Leverage', text: 'Leverage' }
      ]
    };

    const numberOptions = [];
    for (let i = 1; i <= 10; i++) {
      numberOptions.push({ key: i.toString(), text: i.toString() });
    }
    
    return (
      <div>
        <DatePicker
          label="Identification Date"
          value={formData.identificationDate ? new Date(formData.identificationDate) : undefined}
          onSelectDate={(date) => updateFormData('identificationDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <TextField
          label="Description"
          multiline
          rows={3}
          value={formData.description || ''}
          onChange={(_, value) => updateFormData('description', value || '')}
          placeholder="Enter value here"
        />
        
        <TextField
          label="Applicability"
          value={formData.applicability || 'Yes'}
          onChange={(_, value) => updateFormData('applicability', value || '')}
        />
        
        <Dropdown
          label="Associated Goal"
          options={dropdownOptions.associatedGoal}
          selectedKey={formData.associatedGoal}
          onChange={(_, option) => updateFormData('associatedGoal', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Source"
          options={dropdownOptions.source}
          selectedKey={formData.source}
          onChange={(_, option) => updateFormData('source', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Category"
          options={dropdownOptions.category}
          selectedKey={formData.category}
          onChange={(_, option) => updateFormData('category', option?.key)}
          placeholder="Select..."
        />
        
        <TextField
          label="Impact"
          value={formData.impact || ''}
          onChange={(_, value) => updateFormData('impact', value || '')}
          placeholder="Enter value here"
        />
        
        <Dropdown
          label="Priority"
          options={dropdownOptions.priority}
          selectedKey={formData.priority}
          onChange={(_, option) => updateFormData('priority', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Potential Cost"
          options={numberOptions}
          selectedKey={formData.potentialCost?.toString()}
          onChange={(_, option) => {
            const newCost = Number(option?.key);
            updateFormData('potentialCost', newCost);
            // Calculate OpportunityValue immediately
            const benefit = Number(formData.potentialBenefit) || 0;
            updateFormData('opportunityValue', newCost * benefit);
          }}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Potential Benefit"
          options={numberOptions}
          selectedKey={formData.potentialBenefit?.toString()}
          onChange={(_, option) => {
            const newBenefit = Number(option?.key);
            updateFormData('potentialBenefit', newBenefit);
            // Calculate OpportunityValue immediately
            const cost = Number(formData.potentialCost) || 0;
            updateFormData('opportunityValue', cost * newBenefit);
          }}
          placeholder="Select..."
        />
        
        <div className={styles.calculatedField}>
          <Text variant="medium">Opportunity Value (% to 10)</Text>
          <div className={styles.calculatedValue}>
            {formData.opportunityValue || 0}
          </div>
        </div>
        
        <Dropdown
          label="Type of Action"
          options={dropdownOptions.typeOfAction}
          selectedKey={formData.typeOfAction}
          onChange={(_, option) => updateFormData('typeOfAction', option?.key)}
          placeholder="Select..."
        />
        
        <TextField
          label="Action Plan"
          multiline
          rows={3}
          value={formData.actionPlan || ''}
          onChange={(_, value) => updateFormData('actionPlan', value || '')}
          placeholder="Enter value here"
        />
        
        <PeoplePicker
          key={`responsibility-opportunity-${formData.id || 'new'}-${formData.responsibility ? JSON.stringify(formData.responsibility) : 'empty'}`}
          context={{
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient
          }}
          titleText="Responsibility"
          personSelectionLimit={3}
          groupName=""
          showtooltip={true}
          required={false}
          disabled={false}
          onChange={(items: any[]) => {
            console.log('👤 Responsibility (Opportunity) PeoplePicker onChange:', items);
            const selectedUsers = convertPeoplePickerItems(items);
            console.log('👤 Responsibility (Opportunity) converted users:', selectedUsers);
            updateFormData('responsibility', selectedUsers);
          }}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
          defaultSelectedUsers={formData.responsibility ? convertToPickerFormat(formData.responsibility) : []}
        />
        
        <DatePicker
          label="Target Date"
          value={formData.targetDate ? new Date(formData.targetDate) : undefined}
          onSelectDate={(date) => updateFormData('targetDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <DatePicker
          label="Actual Date"
          value={formData.actualDate ? new Date(formData.actualDate) : undefined}
          onSelectDate={(date) => updateFormData('actualDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <Dropdown
          label="Status"
          options={dropdownOptions.status}
          selectedKey={formData.status}
          onChange={(_, option) => updateFormData('status', option?.key)}
          placeholder="Select..."
        />
        
        <TextField
          label="Effectiveness"
          multiline
          rows={3}
          value={formData.effectiveness || ''}
          onChange={(_, value) => updateFormData('effectiveness', value || '')}
          placeholder="Enter value here"
        />
        
        <TextField
          label="Remarks"
          multiline
          rows={3}
          value={formData.remarks || ''}
          onChange={(_, value) => updateFormData('remarks', value || '')}
          placeholder="Enter value here"
        />
      </div>
    );
  }

  const renderRiskForm = (): React.ReactElement => {
    
    const dropdownOptions = {
      associatedGoal: [
        { key: 'BG01', text: 'BG01' },
        { key: 'BG02', text: 'BG02' }
      ],
      source: [
        { key: 'Internal', text: 'Internal' },
        { key: 'External', text: 'External' }
      ],
      category: [
        { key: 'Resource', text: 'Resource' },
        { key: 'Customer', text: 'Customer' },
        { key: 'Information Security', text: 'Information Security' },
        { key: 'Technology', text: 'Technology' },
        { key: 'Business', text: 'Business' },
        { key: 'Process', text: 'Process' },
        { key: 'Others', text: 'Others' }
      ],
      priority: [
        { key: 'High', text: 'High' },
        { key: 'Medium', text: 'Medium' },
        { key: 'Low', text: 'Low' }
      ],
      actionType: [
        { key: 'Mitigation', text: 'Mitigation' },
        { key: 'Contingency', text: 'Contingency' }
      ],
      status: [
        { key: 'Open', text: 'Open' },
        { key: 'Closed', text: 'Closed' },
        { key: 'In Progress', text: 'In Progress' },
        { key: 'Monitoring', text: 'Monitoring' }
      ]
    };

    const numberOptions = [];
    for (let i = 1; i <= 10; i++) {
      numberOptions.push({ key: i.toString(), text: i.toString() });
    }
    
    return (
      <div>
        <DatePicker
          label="Identification Date"
          value={formData.identificationDate ? new Date(formData.identificationDate) : undefined}
          onSelectDate={(date) => updateFormData('identificationDate', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <TextField
          label="Description"
          multiline
          rows={3}
          value={formData.description || ''}
          onChange={(_, value) => updateFormData('description', value || '')}
          placeholder="Enter value here"
        />
        
        <TextField
          label="Applicability"
          value={formData.applicability || 'Yes'}
          onChange={(_, value) => updateFormData('applicability', value || '')}
        />
        
        <Dropdown
          label="Associated Goal"
          options={dropdownOptions.associatedGoal}
          selectedKey={formData.associatedGoal}
          onChange={(_, option) => updateFormData('associatedGoal', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Source"
          options={dropdownOptions.source}
          selectedKey={formData.source}
          onChange={(_, option) => updateFormData('source', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Category"
          options={dropdownOptions.category}
          selectedKey={formData.category}
          onChange={(_, option) => updateFormData('category', option?.key)}
          placeholder="Select..."
        />
        
        <TextField
          label="Impact"
          value={formData.impact || ''}
          onChange={(_, value) => updateFormData('impact', value || '')}
          placeholder="Enter value here"
        />
        
        <Dropdown
          label="Priority"
          options={dropdownOptions.priority}
          selectedKey={formData.priority}
          onChange={(_, option) => updateFormData('priority', option?.key)}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Impact Value (1 to 10)"
          options={numberOptions}
          selectedKey={formData.impactValue?.toString()}
          onChange={(_, option) => {
            const newImpact = Number(option?.key);
            updateFormData('impactValue', newImpact);
            // Calculate RiskExposure immediately
            const probability = Number(formData.probabilityValue) || 0;
            updateFormData('riskExposure', newImpact * probability);
          }}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Probability Value (1 to 10)"
          options={numberOptions}
          selectedKey={formData.probabilityValue?.toString()}
          onChange={(_, option) => {
            const newProbability = Number(option?.key);
            updateFormData('probabilityValue', newProbability);
            // Calculate RiskExposure immediately
            const impact = Number(formData.impactValue) || 0;
            updateFormData('riskExposure', impact * newProbability);
          }}
          placeholder="Select..."
        />
        
        <div className={styles.calculatedField}>
          <Text variant="medium">Risk Exposure</Text>
          <div className={styles.calculatedValue}>
            {formData.riskExposure || 0}
          </div>
        </div>
        
        <div className={styles.actionsSection}>
          <Text variant="medium">Type of Action</Text>
          <div style={{ display: 'flex', gap: '16px', marginBottom: '16px', alignItems: 'center' }}>
            <Checkbox
              label="Mitigation"
              checked={selectedActionTypes.indexOf('Mitigation') !== -1}
              onChange={(_, checked) => handleActionTypeChange('Mitigation', checked || false)}
            />
            <Checkbox
              label="Contingency"
              checked={selectedActionTypes.indexOf('Contingency') !== -1}
              onChange={(_, checked) => handleActionTypeChange('Contingency', checked || false)}
            />
          </div>
          
          {selectedActionTypes.length > 0 && (
            <div style={{ marginTop: '16px', border: '1px solid #e0e0e0', borderRadius: '4px', background: 'white' }}>
              <Pivot>
                {selectedActionTypes.indexOf('Mitigation') !== -1 && mitigationAction && (
                  <PivotItem headerText="Mitigation">
                    <div style={{ padding: '16px', background: '#fefefe' }}>
                      <div className={styles.actionRow}>
                        <div className={styles.actionHeader}>
                          <Text variant="mediumPlus">Mitigation Action</Text>
                        </div>
                        
                        <TextField
                          label="Action Plan"
                          multiline
                          rows={2}
                          value={mitigationAction.plan}
                          onChange={(_, value) => updateActionInType('Mitigation', 'plan', value || '')}
                          placeholder="Enter action plan here"
                        />
                        
                        <PeoplePicker
                          key={`responsibility-mitigation-${formData.id || 'new'}-${mitigationAction.responsibility ? JSON.stringify(mitigationAction.responsibility) : 'empty'}`}
                          context={{
                            absoluteUrl: context.pageContext.web.absoluteUrl,
                            msGraphClientFactory: context.msGraphClientFactory,
                            spHttpClient: context.spHttpClient
                          }}
                          titleText="Responsibility"
                          personSelectionLimit={3}
                          groupName=""
                          showtooltip={true}
                          required={false}
                          disabled={false}
                          onChange={(items: any[]) => {
                            console.log('👤 Responsibility (Mitigation) PeoplePicker onChange:', items);
                            const selectedUsers = convertPeoplePickerItems(items);
                            console.log('👤 Responsibility (Mitigation) converted users:', selectedUsers);
                            updateActionInType('Mitigation', 'responsibility', selectedUsers);
                          }}
                          showHiddenInUI={false}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          defaultSelectedUsers={mitigationAction.responsibility ? convertToPickerFormat(mitigationAction.responsibility) : []}
                        />
                        
                        <DatePicker
                          label="Target Date"
                          value={mitigationAction.targetDate ? new Date(mitigationAction.targetDate) : undefined}
                          onSelectDate={(date) => updateActionInType('Mitigation', 'targetDate', date?.toISOString().split('T')[0] || '')}
                          placeholder="Select a date"
                        />
                        
                        <DatePicker
                          label="Actual Date"
                          value={mitigationAction.actualDate ? new Date(mitigationAction.actualDate) : undefined}
                          onSelectDate={(date) => updateActionInType('Mitigation', 'actualDate', date?.toISOString().split('T')[0] || '')}
                          placeholder="Select a date"
                        />
                        
                        <Dropdown
                          label="Status"
                          options={dropdownOptions.status}
                          selectedKey={mitigationAction.status}
                          onChange={(_, option) => updateActionInType('Mitigation', 'status', option?.key as string)}
                          placeholder="Select..."
                        />
                      </div>
                    </div>
                  </PivotItem>
                )}
                
                {selectedActionTypes.indexOf('Contingency') !== -1 && contingencyAction && (
                  <PivotItem headerText="Contingency">
                    <div style={{ padding: '16px', background: '#fefefe' }}>
                      <div className={styles.actionRow}>
                        <div className={styles.actionHeader}>
                          <Text variant="mediumPlus">Contingency Action</Text>
                        </div>
                        
                        <TextField
                          label="Action Plan"
                          multiline
                          rows={2}
                          value={contingencyAction.plan}
                          onChange={(_, value) => updateActionInType('Contingency', 'plan', value || '')}
                          placeholder="Enter action plan here"
                        />
                        
                        <PeoplePicker
                          key={`responsibility-contingency-${formData.id || 'new'}-${contingencyAction.responsibility ? JSON.stringify(contingencyAction.responsibility) : 'empty'}`}
                          context={{
                            absoluteUrl: context.pageContext.web.absoluteUrl,
                            msGraphClientFactory: context.msGraphClientFactory,
                            spHttpClient: context.spHttpClient
                          }}
                          titleText="Responsibility"
                          personSelectionLimit={3}
                          groupName=""
                          showtooltip={true}
                          required={false}
                          disabled={false}
                          onChange={(items: any[]) => {
                            console.log('👤 Responsibility (Contingency) PeoplePicker onChange:', items);
                            const selectedUsers = convertPeoplePickerItems(items);
                            console.log('👤 Responsibility (Contingency) converted users:', selectedUsers);
                            updateActionInType('Contingency', 'responsibility', selectedUsers);
                          }}
                          showHiddenInUI={false}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          defaultSelectedUsers={contingencyAction.responsibility ? convertToPickerFormat(contingencyAction.responsibility) : []}
                        />
                        
                        <DatePicker
                          label="Target Date"
                          value={contingencyAction.targetDate ? new Date(contingencyAction.targetDate) : undefined}
                          onSelectDate={(date) => updateActionInType('Contingency', 'targetDate', date?.toISOString().split('T')[0] || '')}
                          placeholder="Select a date"
                        />
                        
                        <DatePicker
                          label="Actual Date"
                          value={contingencyAction.actualDate ? new Date(contingencyAction.actualDate) : undefined}
                          onSelectDate={(date) => updateActionInType('Contingency', 'actualDate', date?.toISOString().split('T')[0] || '')}
                          placeholder="Select a date"
                        />
                        
                        <Dropdown
                          label="Status"
                          options={dropdownOptions.status}
                          selectedKey={contingencyAction.status}
                          onChange={(_, option) => updateActionInType('Contingency', 'status', option?.key as string)}
                          placeholder="Select..."
                        />
                      </div>
                    </div>
                  </PivotItem>
                )}
              </Pivot>
            </div>
          )}
        </div>
        
        <TextField
          label="Effectiveness"
          multiline
          rows={3}
          value={formData.effectiveness || ''}
          onChange={(_, value) => updateFormData('effectiveness', value || '')}
          placeholder="Enter value here"
        />
        
        <TextField
          label="Remarks"
          multiline
          rows={3}
          value={formData.remarks || ''}
          onChange={(_, value) => updateFormData('remarks', value || '')}
          placeholder="Enter value here"
        />
      </div>
    );
  }

  const getModalTitle = (): string => {
    const capitalize = (str: string) => str.charAt(0).toUpperCase() + str.slice(1);
    return item ? `Edit ${capitalize(type)}` : `New ${capitalize(type)}`;
  };

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onCancel}
      isBlocking={true}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalContent}>
        <div className={styles.modalHeader}>
          <h2>{getModalTitle()}</h2>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close"
            onClick={onCancel}
            className={styles.closeButton}
          />
        </div>
        
        <div className={styles.modalBody}>
          {renderFormFields()}
        </div>
        
        <div className={styles.modalFooter}>
          <DefaultButton text="Cancel" onClick={onCancel} />
          <PrimaryButton text="Save" onClick={handleSave} />
        </div>
      </div>
    </Modal>
  );
};

export default RaidForm;