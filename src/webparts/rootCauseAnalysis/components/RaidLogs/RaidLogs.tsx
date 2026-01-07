import * as React from 'react';
import { PrimaryButton, Modal, IconButton, Pivot, PivotItem, Spinner } from '@fluentui/react';
import styles from './RaidLogs.module.scss';
import { IRaidLogsProps, IRaidItem, RaidType } from './interfaces/IRaidItem';
import RaidTable from './RaidTables';
import RaidForm from './RaidForms';
import { RaidServiceFactory } from './RaidListService';
import { IExtendedRaidItem } from './interfaces/IRaidService';
import { SUCCESS_MESSAGES, ERROR_MESSAGES } from '../../../../common/Constants';
import { MessageModal, MessageType } from '../ModalPopups';

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
        <h1>RAID Logs</h1>
        <PrimaryButton 
          text="+ Add New" 
          onClick={openNewItemModal}
          className={styles.addButton}
          disabled={loading}
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

export default RaidLogs;