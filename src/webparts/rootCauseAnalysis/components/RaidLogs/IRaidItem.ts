import { WebPartContext } from '@microsoft/sp-webpart-base';

export type RaidType = 'Risk' | 'Opportunity' | 'Issue' | 'Assumption' | 'Dependency';

export interface IPersonPickerUser {
  id: string;
  loginName: string;
  displayName: string;
  email: string;
}

export interface IRaidAction {
  type: string;
  plan: string;
  responsibility: IPersonPickerUser[];
  targetDate: string;
  actualDate: string;
  status: string;
}

export interface IRaidItem {
  id: number;
  type: RaidType;
  
  // Common fields for Risk and Opportunity
  identificationDate?: string;
  description?: string;
  applicability?: string;
  associatedGoal?: string;
  source?: string;
  category?: string;
  impact?: string;
  priority?: string;
  
  // Risk specific fields
  impactValue?: number;
  probabilityValue?: number;
  riskExposure?: number;
  actions?: IRaidAction[];
  
  // Opportunity specific fields
  potentialCost?: number;
  potentialBenefit?: number;
  opportunityValue?: number;
  typeOfAction?: string;
  actionPlan?: string;
  responsibility?: IPersonPickerUser[];
  targetDate?: string;
  actualDate?: string;
  status?: string;
  
  // Issue/Assumption/Dependency specific fields
  details?: string;
  date?: string;
  byWhom?: IPersonPickerUser[];
  implementationActions?: string;
  plannedClosureDate?: string;
  actualClosureDate?: string;
  
  // Common fields
  effectiveness?: string;
  remarks?: string;
}

export interface IRaidLogsProps {
  context: WebPartContext;
}

export interface IRaidLogsState {
  items: IRaidItem[];
  filteredItems: IRaidItem[];
  currentTab: RaidType | 'all';
  showModal: boolean;
  showTypeModal: boolean;
  currentItem: IRaidItem | null;
  editingId: number | null;
  selectedType: RaidType | null;
}