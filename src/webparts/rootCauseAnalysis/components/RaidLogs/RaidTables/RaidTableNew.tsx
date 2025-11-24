import * as React from 'react';
import { DetailsList, IColumn, SelectionMode, IconButton, DetailsRow } from '@fluentui/react';
import styles from './RaidTable.module.scss';
import { RaidType, IPersonPickerUser } from '../interfaces/IRaidItem';
import { IExtendedRaidItem } from '../interfaces/IRaidService';
import { formatDateShort } from '../../../../../common/DateUtils';

export interface IRaidTableProps {
  items: IExtendedRaidItem[];
  currentTab: RaidType;
  onEdit: (item: IExtendedRaidItem) => void;
  onDelete: (item: IExtendedRaidItem) => Promise<void>;
  onViewHistory?: (item: IExtendedRaidItem) => void;
}

const RaidTable: React.FC<IRaidTableProps> = ({ items, currentTab, onEdit, onDelete, onViewHistory }) => {
  const [expandedRows, setExpandedRows] = React.useState<Set<number>>(new Set());

  const toggleRow = (itemId: number): void => {
    setExpandedRows((prevExpanded) => {
      const newExpanded = new Set<number>();
      prevExpanded.forEach(id => newExpanded.add(id));
      
      if (newExpanded.has(itemId)) {
        newExpanded.delete(itemId);
      } else {
        newExpanded.add(itemId);
      }
      return newExpanded;
    });
  };
  
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
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={`${styles.actionButton} ${styles.editButton}`}
              />
              {/* Delete button intentionally hidden per UI requirement */}
            </div>
          )
        },
        {
          key: 'details',
          name: 'Details',
          fieldName: 'details',
          minWidth: 200,
          maxWidth: 250,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.details)
        },
        {
          key: 'date',
          name: 'Date',
          fieldName: 'date',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.date)
        },
        {
          key: 'byWhom',
          name: 'Identified By',
          fieldName: 'byWhom',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.byWhom)
        },
        {
          key: 'implementationActions',
          name: 'Implementation Actions',
          fieldName: 'implementationActions',
          minWidth: 180,
          maxWidth: 220,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.implementationActions, 40)
        },
        {
          key: 'responsibility',
          name: 'Responsibility',
          fieldName: 'responsibility',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.responsibility)
        },
        {
          key: 'plannedClosureDate',
          name: 'Planned Closure Date',
          fieldName: 'plannedClosureDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.plannedClosureDate)
        },
        {
          key: 'actualClosureDate',
          name: 'Actual Closure Date',
          fieldName: 'actualClosureDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.actualClosureDate)
        },
        {
          key: 'remarks',
          name: 'Remarks',
          fieldName: 'remarks',
          minWidth: 150,
          maxWidth: 200,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.remarks, 40)
        },
        
      ];
    } else if (currentTab === 'Opportunity') {
      return [
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={`${styles.actionButton} ${styles.editButton}`}
              />
              <IconButton
                iconProps={{ iconName: 'History' }}
                title="Version History"
                onClick={() => onViewHistory && onViewHistory(item)}
                className={`${styles.actionButton} ${styles.historyButton}`}
              />
              {/* Delete button intentionally hidden per UI requirement */}
            </div>
          )
        },
        {
          key: 'identificationDate',
          name: 'Identification Date',
          fieldName: 'identificationDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.identificationDate)
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.priority || '-'
        },
        {
          key: 'potentialCost',
          name: 'Potential Cost',
          fieldName: 'potentialCost',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
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
          isResizable: true,
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
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => (
            <div>{item.opportunityValue !== undefined ? item.opportunityValue: '-'}</div>
          )
        },
        {
          key: 'leverageActionPlan',
          name: 'Leverage Action Plan',
          fieldName: 'actionPlan',
          minWidth: 180,
          maxWidth: 300,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => (
            <div>{item.actionPlan ? <div className={styles.subText}>{truncate(item.actionPlan, 140)}</div> : '-'}</div>
          )
        },
        {
          key: 'responsibility',
          name: 'Responsibility',
          fieldName: 'responsibility',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatUserArray(item.responsibility)
        },
        {
          key: 'targetDate',
          name: 'Target Date',
          fieldName: 'targetDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.targetDate)
        },
        {
          key: 'actualDate',
          name: 'Actual Date',
          fieldName: 'actualDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.actualDate)
        },
        {
          key: 'status',
          name: 'Status',
          fieldName: 'status',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.status || '-'
        },
        {
          key: 'effectiveness',
          name: 'Effectiveness',
          fieldName: 'effectiveness',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.effectiveness || '-'
        },
        {
          key: 'remarks',
          name: 'Remarks',
          fieldName: 'remarks',
          minWidth: 150,
          maxWidth: 200,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.remarks, 40)
        },
        
      ];
    } else if (currentTab === 'Risk') {
      return [
        {
          key: 'expand',
          name: '',
          minWidth: 40,
          maxWidth: 40,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => {
            const hasActions = item.actions && item.actions.length > 0;
            if (!hasActions) return null;
            const isExpanded = expandedRows.has(item.id);
            return (
              <IconButton
                iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                title={isExpanded ? 'Collapse' : 'Expand'}
                onClick={() => toggleRow(item.id)}
                className={styles.expandButton}
              />
            );
          }
        },
        {
          key: 'actions',
          name: 'Actions',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => (
            <div className={styles.actions}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={() => onEdit(item)}
                className={`${styles.actionButton} ${styles.editButton}`}
              />
              <IconButton
                iconProps={{ iconName: 'History' }}
                title="Version History"
                onClick={() => onViewHistory && onViewHistory(item)}
                className={`${styles.actionButton} ${styles.historyButton}`}
              />
              {/* Delete button intentionally hidden per UI requirement */}
            </div>
          )
        },
        {
          key: 'identificationDate',
          name: 'Identification Date',
          fieldName: 'identificationDate',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => formatDateShort(item.identificationDate)
        },
        {
          key: 'description',
          name: 'Description',
          fieldName: 'description',
          minWidth: 180,
          maxWidth: 220,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.description, 40)
        },
        {
          key: 'associatedGoal',
          name: 'Associated Goal',
          fieldName: 'associatedGoal',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.associatedGoal || '-'
        },
        {
          key: 'source',
          name: 'Source',
          fieldName: 'source',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.source || '-'
        },
        {
          key: 'category',
          name: 'Category',
          fieldName: 'category',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.category || '-'
        },
        {
          key: 'impact',
          name: 'Impact',
          fieldName: 'impact',
          minWidth: 100,
          maxWidth: 130,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.impact || '-'
        },
        {
          key: 'priority',
          name: 'Priority',
          fieldName: 'priority',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.priority || '-'
        },
        {
          key: 'impactValue',
          name: 'Impact Value',
          fieldName: 'impactValue',
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.impactValue?.toString() || '-'
        },
        {
          key: 'probabilityValue',
          name: 'Probability Value',
          fieldName: 'probabilityValue',
          minWidth: 130,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.probabilityValue?.toString() || '-'
        },
        {
          key: 'riskExposure',
          name: 'Risk Exposure',
          fieldName: 'riskExposure',
          minWidth: 110,
          maxWidth: 130,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.riskExposure?.toString() || '-'
        },
        {
          key: 'typeOfAction',
          name: 'Type of Action',
          fieldName: 'typeOfAction',
          minWidth: 150,
          maxWidth: 180,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => getTypeOfAction(item)
        },
        {
          key: 'effectiveness',
          name: 'Effectiveness',
          fieldName: 'effectiveness',
          minWidth: 120,
          maxWidth: 150,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => item.effectiveness || '-'
        },
        {
          key: 'remarks',
          name: 'Remarks',
          fieldName: 'remarks',
          minWidth: 150,
          maxWidth: 200,
          isResizable: true,
          onRender: (item: IExtendedRaidItem) => truncate(item.remarks, 40)
        },
        
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

  // Render Risk table with expandable action details using DetailsList
  if (currentTab === 'Risk') {
    return (
      <div className={styles.tableContainer}>
        <DetailsList
          items={items}
          columns={getColumns()}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          className={styles.detailsList}
          onRenderRow={(props) => {
            if (!props) return null;
            
            const item = props.item as IExtendedRaidItem;
            const isExpanded = expandedRows.has(item.id);
            const hasActions = item.actions && item.actions.length > 0;
            
            return (
              <div key={item.id}>
                <DetailsRow {...props} />
                {isExpanded && hasActions && (
                  <div className={styles.expandedContent}>
                    <div className={styles.actionPlanTable}>
                      <div className={styles.actionPlanHeader}>
                        <div className={styles.actionHeaderCell} style={{ width: '120px' }}>Type of Action</div>
                        <div className={styles.actionHeaderCell} style={{ width: '250px' }}>Action Plan</div>
                        <div className={styles.actionHeaderCell} style={{ width: '120px' }}>Responsibility</div>
                        <div className={styles.actionHeaderCell} style={{ width: '110px' }}>Target Date</div>
                        <div className={styles.actionHeaderCell} style={{ width: '110px' }}>Actual Date</div>
                        <div className={styles.actionHeaderCell} style={{ width: '100px' }}>Status</div>
                      </div>
                      {item.actions && item.actions.map((action: any, index: number) => (
                        <div key={index} className={styles.actionPlanRow}>
                          <div className={styles.actionCell} style={{ width: '120px' }}>{action.type || '-'}</div>
                          <div className={styles.actionCell} style={{ width: '250px' }} title={action.plan}>{truncate(action.plan, 35)}</div>
                          <div className={styles.actionCell} style={{ width: '120px' }}>{formatUserArray(action.responsibility)}</div>
                          <div className={styles.actionCell} style={{ width: '110px' }}>{formatDateShort(action.targetDate)}</div>
                          <div className={styles.actionCell} style={{ width: '110px' }}>{formatDateShort(action.actualDate)}</div>
                          <div className={styles.actionCell} style={{ width: '100px' }}>{action.status || '-'}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          }}
        />
      </div>
    );
  }

  // Default DetailsList for other tabs
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