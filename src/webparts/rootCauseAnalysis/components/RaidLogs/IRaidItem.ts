import { WebPartContext } from '@microsoft/sp-webpart-base';

export type RaidType = 'Risk' | 'Opportunity' | 'Issue' | 'Assumption' | 'Dependency' | 'Constraints';

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
  raidId?: string;
  
  identificationDate?: string;
  description?: string;
  associatedGoal?: string;
  source?: string;
  category?: string;
  impact?: string;
  priority?: string;
  
  impactValue?: number;
  probabilityValue?: number;
  riskExposure?: number;
  typeOfAction?: string;
  actions?: IRaidAction[];
  
  potentialCost?: number;
  potentialBenefit?: number;
  opportunityValue?: number;
  actionPlan?: string;
  responsibility?: IPersonPickerUser[];
  targetDate?: string;
  actualDate?: string;
  status?: string;
  
  details?: string;
  date?: string;
  byWhom?: IPersonPickerUser[];
  implementationActions?: string;
  plannedClosureDate?: string;
  actualClosureDate?: string;
  
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