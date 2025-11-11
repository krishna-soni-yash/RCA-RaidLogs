import * as React from 'react';
import { DetailsList, IColumn, SelectionMode, IconButton } from '@fluentui/react';
import styles from './RaidTable.module.scss';
import { IRaidItem, RaidType, IPersonPickerUser } from '../IRaidItem';
import { IExtendedRaidItem } from '../../../interfaces/IRaidService';

export interface IRaidTableProps {
  items: IExtendedRaidItem[];
  currentTab: RaidType | 'all';
  onEdit: (item: IExtendedRaidItem) => void;
  onDelete: (id: number) => Promise<void>;
}

const RaidTable: React.FC<IRaidTableProps> = ({ items, currentTab, onEdit, onDelete }) => {
  
  const truncate = (str: string | undefined, length: number = 50): string => {
    if (!str) return '-';
    return str.length > length ? str.substring(0, length) + '...' : str;
  };

  const capitalize = (str: string): string => {
    return str.charAt(0).toUpperCase() + str.slice(1);
  };

  const formatUserArray = (users: IPersonPickerUser[] | string | undefined): string => {
    if (!users) return '-';
    
    // Handle backward compatibility with string values
    if (typeof users === 'string') {
      return users;
    }
    
    // Handle array of users
    if (Array.isArray(users) && users.length > 0) {
      return users.map(user => user.displayName).join(', ');
    }
    
    return '-';
  };

  const getColumns = (): IColumn[] => {
    if (currentTab === 'all') {
      return [
        {
          key: 'type',
          name: 'Type',
          fieldName: 'type',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => capitalize(item.type)
        },
        {
          key: 'description',
          name: 'Description/Details',
          fieldName: 'description',
          minWidth: 200,
          maxWidth: 300,
          onRender: (item: IExtendedRaidItem) => truncate(item.description || item.details)
        },
        {
          key: 'date',
          name: 'Date',
          fieldName: 'date',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.date || item.identificationDate || '-'
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
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.priority || '-'
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
                onClick={async () => await onDelete(item.id)}
                className={`${styles.actionButton} ${styles.deleteButton}`}
              />
            </div>
          )
        }
      ];
    } else if (currentTab === 'Issue' || currentTab === 'Assumption' || currentTab === 'Dependency') {
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
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IExtendedRaidItem) => item.date || '-'
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
          minWidth: 140,
          maxWidth: 160,
          onRender: (item: IExtendedRaidItem) => item.plannedClosureDate || '-'
        },
        {
          key: 'actualClosureDate',
          name: 'Actual Closure Date',
          fieldName: 'actualClosureDate',
          minWidth: 140,
          maxWidth: 160,
          onRender: (item: IExtendedRaidItem) => item.actualClosureDate || '-'
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
          onRender: (item: IRaidItem) => (
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
                onClick={() => onDelete(item.id)}
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
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.identificationDate || '-'
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          onRender: (item: IRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          onRender: (item: IRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.priority || '-'
        },
        {
          key: 'potentialCost',
          name: 'Potential Cost',
          fieldName: 'potentialCost',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.potentialCost?.toString() || '-'
        },
        {
          key: 'potentialBenefit',
          name: 'Potential Benefit',
          fieldName: 'potentialBenefit',
          minWidth: 120,
          maxWidth: 140,
          onRender: (item: IRaidItem) => item.potentialBenefit?.toString() || '-'
        },
        {
          key: 'opportunityValue',
          name: 'Opportunity Value',
          fieldName: 'opportunityValue',
          minWidth: 120,
          maxWidth: 140,
          onRender: (item: IRaidItem) => item.opportunityValue?.toString() || '-'
        },
        {
          key: 'typeOfAction',
          name: 'Type of Action',
          fieldName: 'typeOfAction',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.typeOfAction || '-'
        },
        {
          key: 'responsibility',
          name: 'Responsibility',
          fieldName: 'responsibility',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => formatUserArray(item.responsibility)
        },
        {
          key: 'targetDate',
          name: 'Target Date',
          fieldName: 'targetDate',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.targetDate || '-'
        },
        {
          key: 'actualDate',
          name: 'Actual Date',
          fieldName: 'actualDate',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.actualDate || '-'
        },
        {
          key: 'status',
          name: 'Status',
          fieldName: 'status',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.status || '-'
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => (
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
                onClick={() => onDelete(item.id)}
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
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.identificationDate || '-'
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          onRender: (item: IRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          onRender: (item: IRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.priority || '-'
        },
        {
          key: 'impactValue',
          name: 'Impact Value',
          fieldName: 'impactValue',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.impactValue?.toString() || '-'
        },
        {
          key: 'probabilityValue',
          name: 'Probability Value',
          fieldName: 'probabilityValue',
          minWidth: 130,
          maxWidth: 150,
          onRender: (item: IRaidItem) => item.probabilityValue?.toString() || '-'
        },
        {
          key: 'riskExposure',
          name: 'Risk Exposure',
          fieldName: 'riskExposure',
          minWidth: 110,
          maxWidth: 130,
          onRender: (item: IRaidItem) => item.riskExposure?.toString() || '-'
        },
        {
          key: 'actionsCount',
          name: 'Actions Count',
          fieldName: 'actionsCount',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => item.actions ? item.actions.length.toString() : '0'
        },
        {
          key: 'status',
          name: 'Status',
          fieldName: 'status',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => {
            if (item.actions && item.actions.length > 0) {
              return item.actions[0].status || '-';
            }
            return '-';
          }
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          onRender: (item: IRaidItem) => (
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
                onClick={() => onDelete(item.id)}
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
            Click "Add New" to create your first {currentTab === 'all' ? 'RAID' : currentTab} item
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