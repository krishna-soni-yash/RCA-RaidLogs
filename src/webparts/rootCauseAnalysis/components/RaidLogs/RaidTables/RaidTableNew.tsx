import * as React from 'react';
import { DetailsList, IColumn, SelectionMode, IconButton } from '@fluentui/react';
import styles from './RaidTable.module.scss';
import { RaidType, IPersonPickerUser } from '../IRaidItem';
import { IExtendedRaidItem } from '../../../interfaces/IRaidService';
import { formatDateShort } from '../../../constants/DateUtils';

export interface IRaidTableProps {
  items: IExtendedRaidItem[];
  currentTab: RaidType;
  onEdit: (item: IExtendedRaidItem) => void;
  onDelete: (item: IExtendedRaidItem) => Promise<void>;
}

const RaidTable: React.FC<IRaidTableProps> = ({ items, currentTab, onEdit, onDelete }) => {
  
  const truncate = (str: string | undefined, length: number = 50): string => {
    if (!str) return '-';
    return str.length > length ? str.substring(0, length) + '...' : str;
  };

  const formatUserArray = (users: IPersonPickerUser[] | string | undefined): string => {
    if (!users) return '-';
    
    if (typeof users === 'string') {
      return users;
    }
    
    if (Array.isArray(users) && users.length > 0) {
      return users.map(user => user.displayName).join(', ');
    }
    
    return '-';
  };

  const getTypeOfAction = (item: IExtendedRaidItem): string => {
    if (item.type !== 'Risk') return '-';
    
    if (item.actions && item.actions.length > 0) {
      const actionTypes = item.actions
        .map(action => action.type)
        .filter((type, index, self) => type && self.indexOf(type) === index);
      return actionTypes.length > 0 ? actionTypes.join(', ') : '-';
    }
    
    if (item.typeOfAction) {
      return item.typeOfAction;
    }
    
    return '-';
  };

  const getColumns = (): IColumn[] => {
    if (currentTab === 'Issue' || currentTab === 'Assumption' || currentTab === 'Dependency' || currentTab === 'Constraints') {
      return [
        {
          key: 'details',
          name: 'Details',
          fieldName: 'details',
          minWidth: 200,
          maxWidth: 250,
          onRender: (item: IExtendedRaidItem) => truncate(item.details)
        },
        {
          key: 'date',
          name: 'Date',
          fieldName: 'date',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.date)
        },
        {
          key: 'byWhom',
          name: 'By Whom',
          fieldName: 'byWhom',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.byWhom)
        },
        {
          key: 'implementationActions',
          name: 'Implementation Actions',
          fieldName: 'implementationActions',
          minWidth: 180,
          maxWidth: 220,
          onRender: (item: IExtendedRaidItem) => truncate(item.implementationActions, 40)
        },
        {
          key: 'responsibility',
          name: 'Responsibility',
          fieldName: 'responsibility',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.responsibility)
        },
        {
          key: 'plannedClosureDate',
          name: 'Planned Closure Date',
          fieldName: 'plannedClosureDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.plannedClosureDate)
        },
        {
          key: 'actualClosureDate',
          name: 'Actual Closure Date',
          fieldName: 'actualClosureDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.actualClosureDate)
        },
        {
          key: 'remarks',
          name: 'Remarks',
          fieldName: 'remarks',
          minWidth: 150,
          maxWidth: 200,
          onRender: (item: IExtendedRaidItem) => truncate(item.remarks, 40)
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={styles.actionButton}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete"
                onClick={async () => await onDelete(item)}
                className={`${styles.actionButton} ${styles.deleteButton}`}
              />
            </div>
          )
        }
      ];
    } else if (currentTab === 'Opportunity') {
      return [
        {
          key: 'identificationDate',
          name: 'ID Date',
          fieldName: 'identificationDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.identificationDate)
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          onRender: (item: IExtendedRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          onRender: (item: IExtendedRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.priority || '-'
        },
        {
          key: 'potentialCost',
          name: 'Potential Cost',
          fieldName: 'potentialCost',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => {
            if (!item.potentialCost) return '-';
            const costLabels: { [key: number]: string } = {
              1: '1 - No Cost',
              2: '2 - Very Low Cost',
              3: '3 - Low Cost',
              4: '4 - Medium Cost',
              5: '5 - Moderate Cost',
              6: '6 - Medium Cost',
              7: '7 - High Cost',
              8: '8 - Above High Cost',
              9: '9 - Very High Cost',
              10: '10 - Extreme High Cost'
            };
            return costLabels[item.potentialCost] || item.potentialCost.toString();
          }
        },
        {
          key: 'potentialBenefit',
          name: 'Potential Benefit',
          fieldName: 'potentialBenefit',
          minWidth: 120,
          maxWidth: 140,
          onRender: (item: IExtendedRaidItem) => {
            if (!item.potentialBenefit) return '-';
            const benefitLabels: { [key: number]: string } = {
              1: '1 - No Benefits',
              2: '2 - Low Benefits',
              3: '3 - Moderate Benefits',
              4: '4 - Medium Benefits',
              5: '5 - Above Moderate Benefits',
              6: '6 - Moderate Benefits',
              7: '7 - Medium Benefits',
              8: '8 - Above High Benefits',
              9: '9 - High Benefits',
              10: '10 - Significant Benefits'
            };
            return benefitLabels[item.potentialBenefit] || item.potentialBenefit.toString();
          }
        },
        {
          key: 'opportunityValue',
          name: 'Opportunity Value',
          fieldName: 'opportunityValue',
          minWidth: 120,
          maxWidth: 140,
          onRender: (item: IExtendedRaidItem) => item.opportunityValue?.toString() || '-'
        },
        {
          key: 'responsibility',
          name: 'Responsibility',
          fieldName: 'responsibility',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.responsibility)
        },
        {
          key: 'targetDate',
          name: 'Target Date',
          fieldName: 'targetDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.targetDate)
        },
        {
          key: 'actualDate',
          name: 'Actual Date',
          fieldName: 'actualDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.actualDate)
        },
        {
          key: 'status',
          name: 'Status',
          fieldName: 'status',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.status || '-'
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={styles.actionButton}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete"
                onClick={async () => await onDelete(item)}
                className={`${styles.actionButton} ${styles.deleteButton}`}
              />
            </div>
          )
        }
      ];
    } else if (currentTab === 'Risk') {
      return [
        {
          key: 'identificationDate',
          name: 'ID Date',
          fieldName: 'identificationDate',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.identificationDate)
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          onRender: (item: IExtendedRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          onRender: (item: IExtendedRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.priority || '-'
        },
        {
          key: 'impactValue',
          name: 'Impact Value',
          fieldName: 'impactValue',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.impactValue?.toString() || '-'
        },
        {
          key: 'probabilityValue',
          name: 'Probability Value',
          fieldName: 'probabilityValue',
          minWidth: 130,
          maxWidth: 150,
          onRender: (item: IExtendedRaidItem) => item.probabilityValue?.toString() || '-'
        },
        {
          key: 'riskExposure',
          name: 'Risk Exposure',
          fieldName: 'riskExposure',
          minWidth: 110,
          maxWidth: 130,
          onRender: (item: IExtendedRaidItem) => item.riskExposure?.toString() || '-'
        },
        {
          key: 'typeOfAction',
          name: 'Type of Action',
          fieldName: 'typeOfAction',
          minWidth: 150,
          maxWidth: 180,
          onRender: (item: IExtendedRaidItem) => getTypeOfAction(item)
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={styles.actionButton}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete"
                onClick={async () => await onDelete(item)}
                className={`${styles.actionButton} ${styles.deleteButton}`}
              />
            </div>
          )
        }
      ];
    }
    return [];
  };

  if (items.length === 0) {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyStateContent}>
          <p className={styles.emptyStateTitle}>No items found</p>
          <p className={styles.emptyStateSubtext}>
            Click "Add New" to create your first {currentTab} item
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.tableContainer}>
      <DetailsList
        items={items}
        columns={getColumns()}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
        className={styles.detailsList}
      />
    </div>
  );
};

export default RaidTable;