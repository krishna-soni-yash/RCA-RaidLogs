import * as React from 'react';
import { PrimaryButton, Modal, IconButton, Pivot, PivotItem, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import styles from './RaidLogs.module.scss';
import { IRaidLogsProps, IRaidItem, RaidType } from './IRaidItem';
import RaidTable from './RaidTables';
import RaidForm from './RaidForms';
import { RaidServiceFactory } from '../../services/RaidListService';
import { IExtendedRaidItem } from '../../interfaces/IRaidService';
import { SUCCESS_MESSAGES, ERROR_MESSAGES } from '../../constants/Constants';

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
  const [error, setError] = React.useState<string | null>(null);
  const [success, setSuccess] = React.useState<string | null>(null);
  
  const raidService = React.useMemo(() => RaidServiceFactory.getInstance(context), [context]);

  React.useEffect(() => {
    let timeoutId: number;
    if (success) {
      timeoutId = window.setTimeout(() => setSuccess(null), 5000);
    }
    return () => {
      if (timeoutId) clearTimeout(timeoutId);
    };
  }, [success]);

  React.useEffect(() => {
    let timeoutId: number;
    if (error) {
      timeoutId = window.setTimeout(() => setError(null), 10000);
    }
    return () => {
      if (timeoutId) clearTimeout(timeoutId);
    };
  }, [error]);

  const loadRaidItems = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);
      
      const allItems = await raidService.getAllRaidItems();
      setItems(allItems);
    } catch (err) {
      console.error('Error loading RAID items:', err);
      setError(ERROR_MESSAGES.NETWORK_ERROR);
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
        setError('Failed to load Risk item for editing');
        setLoading(false);
      }
    } else {
      setCurrentItem(item);
      setEditingId(item.id);
      setSelectedType(item.type);
      setShowModal(true);
    }
  };

  const deleteItem = async (item: IExtendedRaidItem): Promise<void> => {
    if (confirm('Are you sure you want to delete this item?')) {
      try {
        setLoading(true);
        setError(null);
        
        let success = false;
        
        if (item.type === 'Risk' && item.raidId) {
          console.log('Deleting Risk items with RaidID:', item.raidId);
          success = await raidService.deleteRiskItemsByRaidId(item.raidId);
        } else {
          success = await raidService.deleteRaidItem(item.id);
        }
        
        if (success) {
          setSuccess('Item deleted successfully');
          await loadRaidItems();
        } else {
          setError(ERROR_MESSAGES.DELETE_FAILED);
        }
      } catch (err) {
        console.error('Error deleting item:', err);
        setError(ERROR_MESSAGES.DELETE_FAILED);
      } finally {
        setLoading(false);
      }
    }
  };

  const saveItem = async (item: IRaidItem): Promise<void> => {
    try {
      setLoading(true);
      setError(null);
      
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
            setSuccess(SUCCESS_MESSAGES.ITEM_UPDATED);
            await loadRaidItems();
          } else {
            setError(ERROR_MESSAGES.UPDATE_FAILED);
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
            setSuccess(SUCCESS_MESSAGES.ITEM_CREATED);
            await loadRaidItems();
          } else {
            setError(ERROR_MESSAGES.CREATE_FAILED);
          }
        }
      } else {
        if (editingId) {
          const updatedItem = await raidService.updateRaidItem(editingId, item);
          
          if (updatedItem) {
            setSuccess(SUCCESS_MESSAGES.ITEM_UPDATED);
            await loadRaidItems();
          } else {
            setError(ERROR_MESSAGES.UPDATE_FAILED);
          }
        } else {
          const newItem = await raidService.createRaidItem(item);
          
          if (newItem) {
            setSuccess(SUCCESS_MESSAGES.ITEM_CREATED);
            await loadRaidItems();
          } else {
            setError(ERROR_MESSAGES.CREATE_FAILED);
          }
        }
      }
      
      closeModal();
    } catch (err) {
      console.error('Error saving item:', err);
      setError(editingId ? ERROR_MESSAGES.UPDATE_FAILED : ERROR_MESSAGES.CREATE_FAILED);
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
      
      {/* Success message */}
      {success && (
        <MessageBar
          messageBarType={MessageBarType.success}
          isMultiline={false}
          onDismiss={() => setSuccess(null)}
          dismissButtonAriaLabel="Close"
        >
          {success}
        </MessageBar>
      )}
      
      {/* Error message */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel="Close"
        >
          {error}
        </MessageBar>
      )}

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
      />

      {renderTypeSelectionModal()}
      
      {showModal && selectedType && (
        <RaidForm
          isOpen={showModal}
          type={selectedType}
          item={currentItem}
          onSave={saveItem}
          onCancel={closeModal}
          context={context}
        />
      )}
    </div>
  );
};

export default RaidLogs;