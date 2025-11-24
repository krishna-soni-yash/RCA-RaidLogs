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
import { IRaidItem, RaidType, IRaidAction, IPersonPickerUser } from '../interfaces/IRaidItem';
import { generateRaidId, DROPDOWN_OPTIONS } from '../../../../../common/Constants';
import { RaidListService } from '../RaidListService';

export interface IRaidFormProps {
  isOpen: boolean;
  type: RaidType;
  item: IRaidItem | null;
  onSave: (item: IRaidItem) => Promise<void>;
  onCancel: () => void;
  context: WebPartContext;
  onValidationError?: (message: string) => void;
}

const RaidForm: React.FC<IRaidFormProps> = ({ isOpen, type, item, onSave, onCancel, context, onValidationError }) => {
  const [formData, setFormData] = React.useState<Partial<IRaidItem>>(item ? { ...item } : { type });
  const [selectedActionTypes, setSelectedActionTypes] = React.useState<string[]>([]);
  const [mitigationAction, setMitigationAction] = React.useState<IRaidAction | null>(null);
  const [contingencyAction, setContingencyAction] = React.useState<IRaidAction | null>(null);
  
  // State for dropdown options fetched from SharePoint lists
  const [potentialCostOptions, setPotentialCostOptions] = React.useState<Array<{ key: string; text: string }>>([]);
  const [potentialBenefitOptions, setPotentialBenefitOptions] = React.useState<Array<{ key: string; text: string }>>([]);
  const [probabilityValueOptions, setProbabilityValueOptions] = React.useState<Array<{ key: string; text: string }>>([]);
  const [impactValueOptions, setImpactValueOptions] = React.useState<Array<{ key: string; text: string }>>([]);

  const convertPeoplePickerItems = (items: any[]): IPersonPickerUser[] => {
    return items.map((p: any) => ({
      id: p?.id ?? p?.loginName ?? p?.key ?? '',
      email: p?.secondaryText ?? p?.mail ?? p?.loginName ?? '',
      loginName: p?.loginName ?? '',
      displayName: p?.text ?? p?.primaryText ?? ''
    }));
  };

  const convertToPickerFormat = (users: IPersonPickerUser[] | any): string[] => {
    if (!users) {
      return [];
    }
    
    const userArray = Array.isArray(users) ? users : [users];
    
    const result = userArray.map(user => {
      if (user.email && user.email.indexOf('@') !== -1) {
        return user.email;
      }
      
      if (user.loginName && user.loginName.indexOf('@') !== -1) {
        return user.loginName;
      }
      
      if (user.loginName && user.loginName.indexOf('i:0#.f|membership|') !== -1) {
        return user.loginName.replace('i:0#.f|membership|', '');
      }
      
      if (user.loginName) {
        return user.loginName;
      }
      
      if (user.displayName) {
        return user.displayName;
      }
      
      if (user.id) {
        return user.id;
      }
      
      return null;
    }).filter(item => item !== null) as string[];
    
    return result;
  };
  const serializeUserField = (users: IPersonPickerUser[]): string => {
    if (!users || !Array.isArray(users) || users.length === 0) return '';
    
    // Serialize as "id|email; id|email" format like RCAForm
    return users
      .map((r: any) => `${r.id ?? ''}|${r.email ?? ''}`)
      .join('; ');
  };
  

  React.useEffect(() => {
    setFormData(item ? { ...item } : { type });
    
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
      
      console.log('📝 RaidForm - Initialized actions:', {
        mitigation: mitigation ? 'exists' : 'none',
        contingency: contingency ? 'exists' : 'none',
        selectedTypes: types
      });
    } else {
      setMitigationAction(null);
      setContingencyAction(null);
      setSelectedActionTypes([]);
    }
  }, [item, type]);

  // Fetch dropdown options from SharePoint lists when component mounts
  React.useEffect(() => {
    const loadDropdownOptions = async () => {
      try {
        const raidService = new RaidListService(context);

        // Fetch all dropdown options in parallel
        const [potentialCost, potentialBenefit, probabilityValue, impactValue] = await Promise.all([
          raidService.getPotentialCostOptions(),
          raidService.getPotentialBenefitOptions(),
          raidService.getProbabilityValueOptions(),
          raidService.getImpactValueOptions()
        ]);

        // Update state with fetched options (no fallback to constants)
        setPotentialCostOptions(potentialCost);
        setPotentialBenefitOptions(potentialBenefit);
        setProbabilityValueOptions(probabilityValue);
        setImpactValueOptions(impactValue);
      } catch (error) {
        console.error('Error loading dropdown options:', error);
        // Set empty arrays on error
        setPotentialCostOptions([]);
        setPotentialBenefitOptions([]);
        setProbabilityValueOptions([]);
        setImpactValueOptions([]);
      }
    };

    if (isOpen) {
      loadDropdownOptions();
    }
  }, [isOpen, context]);

  React.useEffect(() => {
    if (type === 'Opportunity') {
      const cost = formData.potentialCost;
      const benefit = formData.potentialBenefit;
      
      // Excel formula: =IF(AND(N5="",O5=""),"",IF((O5-N5)/O5*100<=0,0,(O5-N5)/O5*100))
      let calculatedValue: number | undefined;
      
      if (!cost && !benefit) {
        calculatedValue = undefined;
      } else {
        const costNum = Number(cost) || 0;
        const benefitNum = Number(benefit) || 0;
        
        if (benefitNum === 0) {
          calculatedValue = 0;
        } else {
          const percentage = ((benefitNum - costNum) / benefitNum) * 100;
          calculatedValue = percentage <= 0 ? 0 : percentage;
        }
      }
      
      if (formData.opportunityValue !== calculatedValue) {
        updateFormData('opportunityValue', calculatedValue);
      }
    }
  }, [formData.potentialCost, formData.potentialBenefit, type]);

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
    // Validate mandatory fields
    if (type === 'Issue' || type === 'Assumption' || type === 'Dependency' || type === 'Constraints') {
      if (!formData.details || formData.details.trim() === '') {
        if (onValidationError) {
          onValidationError('Details field is mandatory. Please fill in the details before saving.');
        }
        return;
      }
    }
    
    if (type === 'Opportunity' || type === 'Risk') {
      if (!formData.description || formData.description.trim() === '') {
        if (onValidationError) {
          onValidationError('Description field is mandatory. Please fill in the description before saving.');
        }
        return;
      }
    }

    const itemToSave: any = {
      ...formData,
      type,
      id: formData.id || 0
    };

    if (type === 'Risk') {
      if (selectedActionTypes.length === 0) {
        if (onValidationError) {
          onValidationError('Please select at least one Type of Action (Mitigation or Contingency) for Risk items.');
        }
        return;
      }

      // Generate unique RaidID for new Risk items (not in edit mode)
      if (!formData.id || formData.id === 0) {
        itemToSave.raidId = generateRaidId();
        console.log('Generated new RaidID for Risk item:', itemToSave.raidId);
      } else {
        itemToSave.raidId = formData.raidId;
      }

      const preparedMitigationAction = (selectedActionTypes.indexOf('Mitigation') !== -1 && mitigationAction) ? {
        ...mitigationAction,
        responsibility: mitigationAction.responsibility 
          ? serializeUserField(mitigationAction.responsibility)
          : ''
      } : null;

      const preparedContingencyAction = (selectedActionTypes.indexOf('Contingency') !== -1 && contingencyAction) ? {
        ...contingencyAction,
        responsibility: contingencyAction.responsibility 
          ? serializeUserField(contingencyAction.responsibility)
          : ''
      } : null;

      itemToSave.mitigationAction = preparedMitigationAction;
      itemToSave.contingencyAction = preparedContingencyAction;
      
      console.log('📝 RaidForm - Saving Risk with actions:', {
        mitigationChecked: selectedActionTypes.indexOf('Mitigation') !== -1,
        contingencyChecked: selectedActionTypes.indexOf('Contingency') !== -1,
        mitigationAction: preparedMitigationAction ? 'included' : 'not included',
        contingencyAction: preparedContingencyAction ? 'included' : 'not included',
        mitigationResponsibility: preparedMitigationAction?.responsibility,
        contingencyResponsibility: preparedContingencyAction?.responsibility
      });
    }

    if (formData.responsibility) {
      const convertedResponsibility = serializeUserField(formData.responsibility);
      itemToSave.responsibility = convertedResponsibility;
      console.log('📝 RaidForm - Converted responsibility:', {
        original: formData.responsibility,
        converted: convertedResponsibility
      });
    }
    if (formData.byWhom) {
      const convertedByWhom = serializeUserField(formData.byWhom);
      itemToSave.byWhom = convertedByWhom;
      console.log('📝 RaidForm - Converted byWhom:', {
        original: formData.byWhom,
        converted: convertedByWhom
      });
    }
    
    console.log('📝 RaidForm - Final item to save:', itemToSave);
    await onSave(itemToSave);
  };

  const renderFormFields = (): React.ReactElement => {
    if (type === 'Issue' || type === 'Assumption' || type === 'Dependency') {
      return renderIssueForm();
    } else if (type === 'Constraints') {
      return renderConstraintsForm();
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
          required
        />
        
        <DatePicker
          label="Date"
          value={formData.date ? new Date(formData.date) : undefined}
          onSelectDate={(date) => updateFormData('date', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <PeoplePicker
          key={`byWhom-${type}-${formData.id || 'new'}-${formData.byWhom ? JSON.stringify(formData.byWhom) : 'empty'}`}
          context={{
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient
          }}
          titleText="Identified By"
          personSelectionLimit={3}
          groupName=""
          showtooltip={true}
          required={false}
          disabled={false}
          ensureUser={true}
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
          key={`responsibility-${type}-${formData.id || 'new'}-${formData.responsibility ? JSON.stringify(formData.responsibility) : 'empty'}`}
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
          ensureUser={true}
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

  const renderConstraintsForm = (): React.ReactElement => {
    
    return (
      <div>
        <TextField
          label="Details"
          multiline
          rows={3}
          value={formData.details || ''}
          onChange={(_, value) => updateFormData('details', value || '')}
          placeholder="Enter value here"
          required
        />
        
        <DatePicker
          label="Date"
          value={formData.date ? new Date(formData.date) : undefined}
          onSelectDate={(date) => updateFormData('date', date?.toISOString().split('T')[0] || '')}
          placeholder="Select a date"
        />
        
        <PeoplePicker
          key={`byWhom-${type}-${formData.id || 'new'}-${formData.byWhom ? JSON.stringify(formData.byWhom) : 'empty'}`}
          context={{
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient
          }}
          titleText="Identified By"
          personSelectionLimit={3}
          groupName=""
          showtooltip={true}
          required={false}
          disabled={false}
          ensureUser={true}
          onChange={(items: any[]) => {
            console.log('👤 ByWhom (Constraints) PeoplePicker onChange:', items);
            const selectedUsers = convertPeoplePickerItems(items);
            console.log('👤 ByWhom (Constraints) converted users:', selectedUsers);
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
          key={`responsibility-${type}-${formData.id || 'new'}-${formData.responsibility ? JSON.stringify(formData.responsibility) : 'empty'}`}
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
          ensureUser={true}
          onChange={(items: any[]) => {
            console.log('👤 Responsibility (Constraints) PeoplePicker onChange:', items);
            const selectedUsers = convertPeoplePickerItems(items);
            console.log('👤 Responsibility (Constraints) converted users:', selectedUsers);
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
      associatedGoal: DROPDOWN_OPTIONS.ASSOCIATED_GOAL,
      source: DROPDOWN_OPTIONS.SOURCE,
      category: DROPDOWN_OPTIONS.CATEGORY,
      priority: DROPDOWN_OPTIONS.PRIORITY,
      status: DROPDOWN_OPTIONS.STATUS
    };
    
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
          required
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
          options={potentialCostOptions}
          selectedKey={formData.potentialCost?.toString()}
          onChange={(_, option) => {
            updateFormData('potentialCost', Number(option?.key));
          }}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Potential Benefit"
          options={potentialBenefitOptions}
          selectedKey={formData.potentialBenefit?.toString()}
          onChange={(_, option) => {
            updateFormData('potentialBenefit', Number(option?.key));
          }}
          placeholder="Select..."
        />
        
        <div className={styles.calculatedField}>
          <Text variant="medium">Opportunity Value</Text>
          <div className={styles.calculatedValue}>
            {formData.opportunityValue !== undefined ? formData.opportunityValue.toFixed(2) : '-'}
          </div>
        </div>
        
        <TextField
          label="Leverage Action plan"
          multiline
          rows={3}
          value={formData.actionPlan || ''}
          onChange={(_, value) => updateFormData('actionPlan', value || '')}
          placeholder="Enter value here"
        />
        
        <PeoplePicker
          key={`responsibility-${type}-${formData.id || 'new'}-${formData.responsibility ? JSON.stringify(formData.responsibility) : 'empty'}`}
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
          ensureUser={true}
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
      associatedGoal: DROPDOWN_OPTIONS.ASSOCIATED_GOAL,
      source: DROPDOWN_OPTIONS.SOURCE,
      category: DROPDOWN_OPTIONS.CATEGORY,
      priority: DROPDOWN_OPTIONS.PRIORITY,
      actionType: DROPDOWN_OPTIONS.ACTION_TYPE,
      status: DROPDOWN_OPTIONS.STATUS
    };
    
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
          required
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
          options={impactValueOptions}
          selectedKey={formData.impactValue?.toString()}
          onChange={(_, option) => {
            updateFormData('impactValue', Number(option?.key));
          }}
          placeholder="Select..."
        />
        
        <Dropdown
          label="Probability Value (1 to 10)"
          options={probabilityValueOptions}
          selectedKey={formData.probabilityValue?.toString()}
          onChange={(_, option) => {
            updateFormData('probabilityValue', Number(option?.key));
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
                          ensureUser={true}
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
                          ensureUser={true}
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

// function onValidationError(arg0: string) {
//   throw new Error('Function not implemented.');
// }
