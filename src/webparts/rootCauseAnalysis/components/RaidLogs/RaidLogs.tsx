import * as React from 'react';
import { DefaultButton, PrimaryButton, Modal, IconButton, Pivot, PivotItem, Spinner } from '@fluentui/react';
import styles from './RaidLogs.module.scss';
import { IRaidLogsProps, IRaidItem, RaidType } from './interfaces/IRaidItem';
import RaidTable from './RaidTables';
import RaidForm from './RaidForms';
import { RaidServiceFactory } from './RaidListService';
import { IExtendedRaidItem } from './interfaces/IRaidService';
import { SUCCESS_MESSAGES, ERROR_MESSAGES } from '../../../../common/Constants';
import { formatDateShort } from '../../../../common/DateUtils';
import { MessageModal, MessageType } from '../ModalPopups';
import { exportRowsToExcel } from '../../../../common/excelExport';

const RaidLogs: React.FC<IRaidLogsProps> = ({ context }) => {
  const [items, setItems] = React.useState<IExtendedRaidItem[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IExtendedRaidItem[]>([]);
  const [currentTab, setCurrentTab] = React.useState<RaidType>('Risk');
  const [showModal, setShowModal] = React.useState<boolean>(false);
  const [showTypeModal, setShowTypeModal] = React.useState<boolean>(false);
  const [currentItem, setCurrentItem] = React.useState<IExtendedRaidItem | null>(null);
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [selectedType, setSelectedType] = React.useState<RaidType | null>(null);
  const [loading, setLoading] = React.useState<boolean>(false);
  
  // Message modal state
  const [showMessageModal, setShowMessageModal] = React.useState<boolean>(false);
  const [messageText, setMessageText] = React.useState<string>('');
  const [messageType, setMessageType] = React.useState<MessageType>('info');
  
  const raidService = React.useMemo(() => RaidServiceFactory.getInstance(context), [context]);

  // History modal state
  const [showHistoryModal, setShowHistoryModal] = React.useState<boolean>(false);
  const [historyVersions, setHistoryVersions] = React.useState<any[]>([]);
  const [historyItemTitle, setHistoryItemTitle] = React.useState<string>('');

  const showMessage = (message: string, type: MessageType): void => {
    setMessageText(message);
    setMessageType(type);
    setShowMessageModal(true);
  };

  const handleDismissMessage = (): void => {
    setShowMessageModal(false);
  };

  const handleValidationError = (message: string): void => {
    showMessage(message, 'warning');
  };

  const loadRaidItems = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      
      const allItems = await raidService.getAllRaidItems();
      setItems(allItems);
    } catch (err) {
      console.error('Error loading RAID items:', err);
      showMessage(ERROR_MESSAGES.NETWORK_ERROR, 'error');
    } finally {
      setLoading(false);
    }
  }, [raidService]);

  const filterItems = React.useCallback((): void => {
    let filtered = items.filter(item => item.type === currentTab);
    
    if (currentTab === 'Risk') {
      const groupedItems: IExtendedRaidItem[] = [];
      const processedRaidIds = new Set<string>();
      
      filtered.forEach(item => {
        if (item.type === 'Risk' && item.raidId) {
          const raidId = item.raidId;
          if (!processedRaidIds.has(raidId)) {
            const relatedItems = filtered.filter(i => i.raidId === raidId);
            
            if (!item.actions || item.actions.length === 0) {
              const actions: any[] = [];
              relatedItems.forEach(relatedItem => {
                if (relatedItem.typeOfAction) {
                  actions.push({
                    type: relatedItem.typeOfAction,
                    plan: relatedItem.actionPlan || '',
                    responsibility: relatedItem.responsibility || [],
                    targetDate: relatedItem.targetDate || '',
                    actualDate: relatedItem.actualDate || '',
                    status: relatedItem.status || ''
                  });
                }
              });
              item = { ...item, actions };
            }
            
            groupedItems.push(item);
            processedRaidIds.add(raidId);
          }
        } else {
          groupedItems.push(item);
        }
      });
      
      filtered = groupedItems;
    }
    
    setFilteredItems(filtered);
  }, [items, currentTab]);

  React.useEffect(() => {
    loadRaidItems();
  }, [loadRaidItems]);

  React.useEffect(() => {
    filterItems();
  }, [filterItems]);

  // If a RaidlogId (or variants) query parameter is present, open that item in the edit modal
  React.useEffect(() => {
    const tryOpenFromQuery = async (): Promise<void> => {
      try {
        const params = new URLSearchParams(window.location.search);
        const raw = params.get('RaidlogId') || params.get('RaidLogId') || params.get('raidlogid') || params.get('RAIDId') || params.get('worklogId');
        if (!raw) return;

        // Try to find item by SP Id first (use explicit loops to avoid lib target issues)
        const value = raw;
        let found: IExtendedRaidItem | undefined = undefined;

        for (let idx = 0; idx < items.length; idx++) {
          const it = items[idx];
          if (String(it.id) === value) {
            found = it;
            break;
          }
        }

        // If not found, try matching raidId (for Risk groups)
        if (!found) {
          for (let idx = 0; idx < items.length; idx++) {
            const it = items[idx];
            if (it.raidId === value) {
              found = it;
              break;
            }
          }
        }

        // If still not found, try fetching by id from service (in case items not yet loaded)
        if (!found && !isNaN(Number(value))) {
          try {
            const fetched = await raidService.getRaidItemById(Number(value));
            if (fetched) {
              found = fetched as IExtendedRaidItem;
            }
          } catch (e) {
            // ignore
          }
        }

        if (found) {
          await editItem(found);

          // Remove query params so modal doesn't reopen on refresh
          const newUrl = new URL(window.location.href);
          newUrl.searchParams.delete('RaidlogId');
          newUrl.searchParams.delete('RaidLogId');
          newUrl.searchParams.delete('raidlogid');
          newUrl.searchParams.delete('RAIDId');
          newUrl.searchParams.delete('worklogId');
          window.history.replaceState(null, '', newUrl.toString());
        }
      } catch (err) {
        console.error('Error opening item from query param:', err);
      }
    };

    if (items && items.length > 0) {
      void tryOpenFromQuery();
    }
  }, [items]);

  const handleTabChange = (item?: PivotItem): void => {
    if (item) {
      const newTab = (item.props.itemKey as RaidType) || 'Risk';
      setCurrentTab(newTab);
    }
  };

  const openNewItemModal = (): void => {
    setShowTypeModal(true);
    setCurrentItem(null);
    setEditingId(null);
  };

  const closeTypeModal = (): void => {
    setShowTypeModal(false);
  };

  const selectType = (type: RaidType): void => {
    setSelectedType(type);
    setShowTypeModal(false);
    setShowModal(true);
  };

  const closeModal = (): void => {
    setShowModal(false);
    setCurrentItem(null);
    setEditingId(null);
    setSelectedType(null);
  };

  const editItem = async (item: IExtendedRaidItem): Promise<void> => {
    // Special handling for Risk type items with RaidID
    if (item.type === 'Risk' && item.raidId) {
      try {
        setLoading(true);
        const riskItems = await raidService.getRiskItemsByRaidId(item.raidId);
        
        if (riskItems && riskItems.length > 0) {
          const baseItem = riskItems[0];
          
          const actions: any[] = [];
          
          riskItems.forEach(riskItem => {
            if (riskItem.typeOfAction === 'Mitigation') {
              actions.push({
                type: 'Mitigation',
                plan: riskItem.actionPlan || '',
                responsibility: riskItem.responsibility || [],
                targetDate: riskItem.targetDate || '',
                actualDate: riskItem.actualDate || '',
                status: riskItem.status || ''
              });
            } else if (riskItem.typeOfAction === 'Contingency') {
              actions.push({
                type: 'Contingency',
                plan: riskItem.actionPlan || '',
                responsibility: riskItem.responsibility || [],
                targetDate: riskItem.targetDate || '',
                actualDate: riskItem.actualDate || '',
                status: riskItem.status || ''
              });
            }
          });
          
          const compositeItem: any = {
            ...baseItem,
            actions
          };
          
          setCurrentItem(compositeItem);
          setEditingId(baseItem.id);
          setSelectedType(baseItem.type);
          setShowModal(true);
        }
        setLoading(false);
      } catch (err) {
        console.error('Error loading Risk item for edit:', err);
        showMessage('Failed to load Risk item for editing', 'error');
        setLoading(false);
      }
    } else {
      setCurrentItem(item);
      setEditingId(item.id);
      setSelectedType(item.type);
      setShowModal(true);
    }
  };

  const viewHistory = async (item: IExtendedRaidItem): Promise<void> => {
    if (!item || !item.id) return;
    try {
      setLoading(true);
      // If this is a Risk item that shares a RAIDId with other items (mitigation/contingency),
      // fetch version history for both SharePoint list items and combine them so the modal
      // shows history across the two related items.
      let combinedVersions: any[] = [];
      if (item.type === 'Risk' && item.raidId) {
        const riskItems = await raidService.getRiskItemsByRaidId(item.raidId);
        if (riskItems && riskItems.length > 1) {
          // Fetch versions for each related item (limit to first two items that share RAIDId)
          const toFetch = riskItems.slice(0, 2);
          const versionsPerItem = await Promise.all(
            toFetch.map(async (ri) => {
              try {
                const v = await raidService.getVersionHistory(ri.id);
                // attach a small marker so we know which SP item these versions belong to
                return (v || []).map((ver: any) => ({ ...ver, __sourceItemId: ri.id, __sourceTypeOfAction: ri.typeOfAction }));
              } catch (e) {
                return [];
              }
            })
          );

          combinedVersions = ([] as any[]).concat(...versionsPerItem);
        } else {
          const versions = await raidService.getVersionHistory(item.id);
          combinedVersions = versions || [];
        }
      } else {
        const versions = await raidService.getVersionHistory(item.id);
        combinedVersions = versions || [];
      }

      setHistoryVersions(combinedVersions);
      setHistoryItemTitle(item.description || `Item ${item.id}`);
      setShowHistoryModal(true);
    } catch (err) {
      console.error('Error fetching version history:', err);
      showMessage(ERROR_MESSAGES.NETWORK_ERROR, 'error');
    } finally {
      setLoading(false);
    }
  };

  const deleteItem = async (item: IExtendedRaidItem): Promise<void> => {
    if (confirm('Are you sure you want to delete this item?')) {
      try {
        setLoading(true);
        
        let success = false;
        
        if (item.type === 'Risk' && item.raidId) {
          console.log('Deleting Risk items with RaidID:', item.raidId);
          success = await raidService.deleteRiskItemsByRaidId(item.raidId);
        } else {
          success = await raidService.deleteRaidItem(item.id);
        }
        
        if (success) {
          showMessage('Item deleted successfully', 'success');
          await loadRaidItems();
        } else {
          showMessage(ERROR_MESSAGES.DELETE_FAILED, 'error');
        }
      } catch (err) {
        console.error('Error deleting item:', err);
        showMessage(ERROR_MESSAGES.DELETE_FAILED, 'error');
      } finally {
        setLoading(false);
      }
    }
  };

  const handleExportRaidTab = (): void => {
    const rows = buildRaidExportRows(currentTab, filteredItems);
    exportRowsToExcel({
      rows,
      headers: RAID_EXPORT_HEADERS[currentTab],
      sheetName: currentTab,
      fileName: `${currentTab}RAIDLogs`
    });
  };

  const saveItem = async (item: IRaidItem): Promise<void> => {
    try {
      setLoading(true);
      
      // Special handling for Risk type items
      if (item.type === 'Risk') {
        const mitigationAction = (item as any).mitigationAction || null;
        const contingencyAction = (item as any).contingencyAction || null;
        
        if (editingId && item.raidId) {
          console.log('Updating Risk items with RaidID:', item.raidId);
          
          const { mitigationAction: _, contingencyAction: __, ...itemWithoutActions } = item as any;
          
          const success = await raidService.updateRiskItemsByRaidId(
            item.raidId,
            itemWithoutActions,
            mitigationAction,
            contingencyAction
          );
          
          if (success) {
            showMessage(SUCCESS_MESSAGES.ITEM_UPDATED, 'success');
            await loadRaidItems();
          } else {
            showMessage(ERROR_MESSAGES.UPDATE_FAILED, 'error');
          }
        } else {
          console.log('Creating new Risk items with RaidID:', item.raidId);
          
          const { mitigationAction: _, contingencyAction: __, ...itemWithoutActions } = item as any;
          
          const createdItems = await raidService.createRiskItemWithActions(
            itemWithoutActions,
            mitigationAction,
            contingencyAction
          );
          
          if (createdItems && createdItems.length > 0) {
            showMessage(SUCCESS_MESSAGES.ITEM_CREATED, 'success');
            await loadRaidItems();
          } else {
            showMessage(ERROR_MESSAGES.CREATE_FAILED, 'error');
          }
        }
      } else {
        if (editingId) {
          const updatedItem = await raidService.updateRaidItem(editingId, item);
          
          if (updatedItem) {
            showMessage(SUCCESS_MESSAGES.ITEM_UPDATED, 'success');
            await loadRaidItems();
          } else {
            showMessage(ERROR_MESSAGES.UPDATE_FAILED, 'error');
          }
        } else {
          const newItem = await raidService.createRaidItem(item);
          
          if (newItem) {
            showMessage(SUCCESS_MESSAGES.ITEM_CREATED, 'success');
            await loadRaidItems();
          } else {
            showMessage(ERROR_MESSAGES.CREATE_FAILED, 'error');
          }
        }
      }
      
      closeModal();
    } catch (err) {
      console.error('Error saving item:', err);
      showMessage(editingId ? ERROR_MESSAGES.UPDATE_FAILED : ERROR_MESSAGES.CREATE_FAILED, 'error');
    } finally {
      setLoading(false);
    }
  };

  const renderTypeSelectionModal = (): React.ReactElement => {
    return (
      <Modal
        isOpen={showTypeModal}
        onDismiss={closeTypeModal}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.modalContent}>
          <div className={styles.modalHeader}>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Close"
              onClick={closeTypeModal}
              className={styles.closeButton}
            />
          </div>
          <div className={styles.modalBody}>
            <div className={styles.typeSelector}>
              <div className={styles.typeCard} onClick={() => selectType('Risk')}>
                <h3>Risk</h3>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Opportunity')}>
                <h3>Opportunity</h3>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Issue')}>
                <h3>Issue</h3>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Assumption')}>
                <h3>Assumption</h3>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Dependency')}>
                <h3>Dependency</h3>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Constraints')}>
                <h3>Constraints</h3>
              </div>
            </div>
          </div>
        </div>
      </Modal>
    );
  };

  return (
    <div className={styles.raidLogs}>
      <div className={styles.header}>
        <PrimaryButton 
          text="+ Add New" 
          onClick={openNewItemModal}
          className={styles.addButton}
          disabled={loading}
        />
        <DefaultButton
          text="Export"
          iconProps={{ iconName: 'Download' }}
          onClick={handleExportRaidTab}
          disabled={loading || filteredItems.length === 0}
        />
      </div>
      
      {/* Loading indicator */}
      {loading && (
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <Spinner label="Loading RAID items..." />
        </div>
      )}
      
      {/* Message Modal */}
      <MessageModal
        isOpen={showMessageModal}
        message={messageText}
        type={messageType}
        onDismiss={handleDismissMessage}
      />

      <div className={styles.tabs}>
        <Pivot
          selectedKey={currentTab}
          onLinkClick={handleTabChange}
          className={styles.pivot}
        >
          <PivotItem headerText="Risk" itemKey="Risk" />
          <PivotItem headerText="Opportunity" itemKey="Opportunity" />
          <PivotItem headerText="Issue" itemKey="Issue" />
          <PivotItem headerText="Assumption" itemKey="Assumption" />
          <PivotItem headerText="Dependency" itemKey="Dependency" />
          <PivotItem headerText="Constraints" itemKey="Constraints" />
        </Pivot>
      </div>

      <RaidTable 
        items={filteredItems}
        currentTab={currentTab}
        onEdit={editItem}
        onDelete={deleteItem}
        onViewHistory={viewHistory}
      />

      {/* Version History Modal */}
      <Modal
        isOpen={showHistoryModal}
        onDismiss={() => setShowHistoryModal(false)}
        isBlocking={false}
        containerClassName={styles.historyModalContainer}
      >
        <div className={styles.historyModalContent}>
          <div className={styles.historyModalHeader}>
            <h2>Version History - {historyItemTitle}</h2>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Close"
              onClick={() => setShowHistoryModal(false)}
              className={styles.closeButton}
            />
          </div>
          <div className={styles.historyModalBody}>
            {historyVersions && historyVersions.length > 0 ? (
              <>
                <div className={styles.versionSuccessMessage}>
                  <i className={`ms-Icon ms-Icon--CompletedSolid ${styles.successIcon}`} />
                  <span>Found {historyVersions.length} version(s) for this item</span>
                </div>
                <div className={styles.versionList}>
                  {historyVersions.map((version: any, index: number) => {
                    const isCurrentVersion = index === 0;
                    const versionLabel = version.VersionLabel || `${version.VersionId || version.Id || index + 1}.0`;
                    const modifiedDate = version.Modified || version.Created || '';
                    const editorName = version.Editor?.Title || version.Editor?.Name || version.Author?.Title || 'Unknown';
                    const editorEmail = version.Editor?.EMail || version.Editor?.Email || '';
                    const modifiedBy = editorEmail ? `(${editorEmail})` : editorName;
                    
                    return (
                      <div 
                        key={version.VersionId || version.Id || index} 
                        className={`${styles.versionCard} ${isCurrentVersion ? styles.currentVersion : ''}`}
                      >
                        <div className={styles.versionHeader}>
                          <div className={styles.versionTitle}>
                            <span className={styles.versionNumber}>Version {versionLabel}</span>
                            {isCurrentVersion && <span className={styles.currentBadge}>(Current)</span>}
                          </div>
                          <div className={styles.versionDate}>
                            {modifiedDate ? new Date(modifiedDate).toLocaleString('en-GB', { 
                              day: '2-digit', 
                              month: '2-digit', 
                              year: 'numeric', 
                              hour: '2-digit', 
                              minute: '2-digit',
                              hour12: true 
                            }) : ''}
                          </div>
                        </div>
                        <div className={styles.versionDetails}>
                          <div className={styles.versionField}>
                            <span className={styles.fieldLabel}>Modified by:</span>
                            <span className={styles.fieldValue}>{modifiedBy}</span>
                          </div>
                          {version.RiskDescription && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Description:</span>
                              <span className={styles.fieldValue}>{version.RiskDescription}</span>
                            </div>
                          )}
                          {version.RiskStatus && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Status:</span>
                              <span className={styles.fieldValue}>{version.RiskStatus}</span>
                            </div>
                          )}
                          {version.Remarks && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Remarks:</span>
                              <span className={styles.fieldValue}>{version.Remarks}</span>
                            </div>
                          )}
                          {version.ActionPlan && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Action Plan:</span>
                              <span className={styles.fieldValue}>{version.ActionPlan}</span>
                            </div>
                          )}
                          {version.Impact && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Impact:</span>
                              <span className={styles.fieldValue}>{version.Impact}</span>
                            </div>
                          )}
                          {version.RiskPriority && (
                            <div className={styles.versionField}>
                              <span className={styles.fieldLabel}>Priority:</span>
                              <span className={styles.fieldValue}>{version.RiskPriority}</span>
                            </div>
                          )}
                          {/* Highlighted Type of Actions block with Action Plan, Responsibility, Target/Actual Date and Status */}
                          {(version.TypeOfAction || version.__sourceTypeOfAction || version.ActionPlan) && (
                            <div style={{ background: '#f5fbf7', padding: 10, borderRadius: 6, marginTop: 8 }}>
                              <div style={{ fontWeight: 600, marginBottom: 6 }}>Type of Actions: <span style={{ fontWeight: 700, color: '#0b6a4a' }}>{version.TypeOfAction || version.__sourceTypeOfAction || '-'}</span></div>
                              <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start', flexWrap: 'wrap' }}>
                                <div style={{ minWidth: 180 }}>
                                  <div className={styles.fieldLabel}>Action Plan</div>
                                  <div className={styles.fieldValue}>{version.ActionPlan || '-'}</div>
                                </div>
                                <div style={{ minWidth: 160 }}>
                                  <div className={styles.fieldLabel}>Responsibility</div>
                                  <div className={styles.fieldValue}>{(version.Responsibility && (version.Responsibility.Title || (Array.isArray(version.Responsibility) ? version.Responsibility.map((r: any) => r.Title || r.displayName).join(', ') : version.Responsibility))) || '-'}</div>
                                </div>
                                <div style={{ minWidth: 120 }}>
                                  <div className={styles.fieldLabel}>Target Date</div>
                                  <div className={styles.fieldValue}>{version.TargetDate ? new Date(version.TargetDate).toLocaleDateString('en-GB') : '-'}</div>
                                </div>
                                <div style={{ minWidth: 120 }}>
                                  <div className={styles.fieldLabel}>Actual Date</div>
                                  <div className={styles.fieldValue}>{version.ActualDate ? new Date(version.ActualDate).toLocaleDateString('en-GB') : '-'}</div>
                                </div>
                                <div style={{ minWidth: 100 }}>
                                  <div className={styles.fieldLabel}>Status</div>
                                  <div className={styles.fieldValue}>{version.RiskStatus || version.Status || '-'}</div>
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            ) : (
              <div className={styles.emptyVersionState}>No version history available</div>
            )}
          </div>
        </div>
      </Modal>

      {renderTypeSelectionModal()}
      
      {showModal && selectedType && (
        <RaidForm
          isOpen={showModal}
          type={selectedType}
          item={currentItem}
          onSave={saveItem}
          onCancel={closeModal}
          context={context}
          onValidationError={handleValidationError}
        />
      )}
    </div>
  );
};

const RAID_EXPORT_HEADERS: Record<RaidType, string[]> = {
  Risk: [
    'Identification Date',
    'Description',
    'Associated Goal',
    'Source',
    'Category',
    'Impact',
    'Priority',
    'Impact Value',
    'Probability Value',
    'Risk Exposure',
    'Type of Action',
    'Action Plan',
    'Responsibility',
    'Target Date',
    'Actual Date',
    'Status',
    'Effectiveness',
    'Remarks'
  ],
  Opportunity: [
    'Identification Date',
    'Description',
    'Associated Goal',
    'Source',
    'Category',
    'Impact',
    'Priority',
    'Potential Cost',
    'Potential Benefit',
    'Opportunity Value',
    'Leverage Action Plan',
    'Responsibility',
    'Target Date',
    'Actual Date',
    'Status',
    'Effectiveness',
    'Remarks'
  ],
  Issue: [
    'Details',
    'Date',
    'Identified By',
    'Implementation Actions',
    'Planned Closure Date',
    'Actual Closure Date',
    'Responsibility',
    'Remarks'
  ],
  Assumption: [
    'Details',
    'Date',
    'Identified By',
    'Implementation Actions',
    'Planned Closure Date',
    'Actual Closure Date',
    'Responsibility',
    'Remarks'
  ],
  Dependency: [
    'Details',
    'Date',
    'Identified By',
    'Implementation Actions',
    'Planned Closure Date',
    'Actual Closure Date',
    'Responsibility',
    'Remarks'
  ],
  Constraints: [
    'Details',
    'Date',
    'Identified By',
    'Implementation Actions',
    'Planned Closure Date',
    'Actual Closure Date',
    'Responsibility',
    'Remarks'
  ]
};

const COST_LABELS: Record<number, string> = {
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

const BENEFIT_LABELS: Record<number, string> = {
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

const buildRaidExportRows = (type: RaidType, items: IExtendedRaidItem[]): Record<string, any>[] => {
  switch (type) {
    case 'Risk': {
      const rows: Record<string, any>[] = [];
      items.forEach(item => {
        const itemRows = mapRiskItem(item);
        itemRows.forEach(row => rows.push(row));
      });
      return rows;
    }
    case 'Opportunity':
      return items.map(mapOpportunityItem);
    case 'Issue':
    case 'Assumption':
    case 'Dependency':
    case 'Constraints':
      return items.map(mapIssueLikeItem);
    default:
      return [];
  }
};

const mapRiskItem = (item: IExtendedRaidItem): Record<string, any>[] => {
  // Get all action types and create Type of Action string
  const actionTypes: string[] = [];
  if (item.actions && item.actions.length > 0) {
    item.actions.forEach(action => {
      if (action.type && actionTypes.indexOf(action.type) === -1) {
        actionTypes.push(action.type);
      }
    });
  }
  const typeOfActionStr = actionTypes.sort().join('; ');

  // Base row data (without action-specific fields)
  const baseRow = {
    'Identification Date': formatDateShort(item.identificationDate),
    'Description': item.description ?? '',
    'Associated Goal': item.associatedGoal ?? '',
    'Source': item.source ?? '',
    'Category': item.category ?? '',
    'Impact': item.impact ?? '',
    'Priority': item.priority ?? '',
    'Impact Value': item.impactValue ?? '',
    'Probability Value': item.probabilityValue ?? '',
    'Risk Exposure': item.riskExposure ?? '',
    'Type of Action': typeOfActionStr,
    'Effectiveness': item.effectiveness ?? '',
    'Remarks': item.remarks ?? ''
  };

  // If no actions or only one action, return single row
  if (!item.actions || item.actions.length === 0) {
    return [{
      ...baseRow,
      'Action Plan': '',
      'Responsibility': '',
      'Target Date': '',
      'Actual Date': '',
      'Status': ''
    }];
  }

  // Check if both Mitigation and Contingency exist
  const hasMitigation = item.actions.some(a => a.type === 'Mitigation');
  const hasContingency = item.actions.some(a => a.type === 'Contingency');

  // If both types exist, create separate rows for each action
  if (hasMitigation && hasContingency) {
    return item.actions.map(action => ({
      ...baseRow,
      'Type of Action': action.type ?? '',
      'Action Plan': action.plan ?? '',
      'Responsibility': formatPersonValue(action.responsibility),
      'Target Date': formatDateShort(action.targetDate),
      'Actual Date': formatDateShort(action.actualDate),
      'Status': action.status ?? ''
    }));
  }

  // If only one type, create single row with action details
  const action = item.actions[0];
  return [{
    ...baseRow,
    'Action Plan': action.plan ?? '',
    'Responsibility': formatPersonValue(action.responsibility),
    'Target Date': formatDateShort(action.targetDate),
    'Actual Date': formatDateShort(action.actualDate),
    'Status': action.status ?? ''
  }];
};

const mapOpportunityItem = (item: IExtendedRaidItem): Record<string, any> => ({
  'Identification Date': formatDateShort(item.identificationDate),
  'Description': item.description ?? '',
  'Associated Goal': item.associatedGoal ?? '',
  'Source': item.source ?? '',
  'Category': item.category ?? '',
  'Impact': item.impact ?? '',
  'Priority': item.priority ?? '',
  'Potential Cost': formatPotentialCost(item.potentialCost),
  'Potential Benefit': formatPotentialBenefit(item.potentialBenefit),
  'Opportunity Value': item.opportunityValue ?? '',
  'Leverage Action Plan': item.actionPlan ?? '',
  'Responsibility': formatPersonValue(item.responsibility),
  'Target Date': formatDateShort(item.targetDate),
  'Actual Date': formatDateShort(item.actualDate),
  'Status': item.status ?? '',
  'Effectiveness': item.effectiveness ?? '',
  'Remarks': item.remarks ?? ''
});

const mapIssueLikeItem = (item: IExtendedRaidItem): Record<string, any> => ({
  'Details': item.details ?? '',
  'Date': formatDateShort(item.date),
  'Identified By': formatPersonValue(item.byWhom),
  'Implementation Actions': item.implementationActions ?? '',
  'Planned Closure Date': formatDateShort(item.plannedClosureDate),
  'Actual Closure Date': formatDateShort(item.actualClosureDate),
  'Responsibility': formatPersonValue(item.responsibility),
  'Remarks': item.remarks ?? ''
});

const formatPersonValue = (value: any): string => {
  if (!value) return '';
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (!entry) return '';
        if (typeof entry === 'string') {
          return entry;
        }
        return entry.displayName || entry.Title || entry.text || entry.EMail || entry.email || '';
      })
      .filter(Boolean)
      .join(', ');
  }
  return '';
};

const formatPotentialCost = (value?: number): string => {
  if (value === undefined || value === null) return '';
  return COST_LABELS[value] || String(value);
};

const formatPotentialBenefit = (value?: number): string => {
  if (value === undefined || value === null) return '';
  return BENEFIT_LABELS[value] || String(value);
};

export default RaidLogs;
