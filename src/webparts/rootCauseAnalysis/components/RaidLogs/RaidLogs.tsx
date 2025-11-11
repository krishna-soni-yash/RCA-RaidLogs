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
  const [currentTab, setCurrentTab] = React.useState<RaidType | 'all'>('all');
  const [showModal, setShowModal] = React.useState<boolean>(false);
  const [showTypeModal, setShowTypeModal] = React.useState<boolean>(false);
  const [currentItem, setCurrentItem] = React.useState<IExtendedRaidItem | null>(null);
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [selectedType, setSelectedType] = React.useState<RaidType | null>(null);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [success, setSuccess] = React.useState<string | null>(null);
  
  // Initialize RAID service
  const raidService = React.useMemo(() => RaidServiceFactory.getInstance(context), [context]);

  // Clear messages after timeout
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

  // Load data from SharePoint on component mount
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
    const filtered = currentTab === 'all' ? items : items.filter(item => item.type === currentTab);
    setFilteredItems(filtered);
  }, [items, currentTab]);

  // Load data on component mount
  React.useEffect(() => {
    loadRaidItems();
  }, [loadRaidItems]);

  React.useEffect(() => {
    filterItems();
  }, [filterItems]);

  const handleTabChange = (item?: PivotItem): void => {
    if (item) {
      const newTab = (item.props.itemKey as RaidType | 'all') || 'all';
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

  const editItem = (item: IExtendedRaidItem): void => {
    setCurrentItem(item);
    setEditingId(item.id);
    setSelectedType(item.type);
    setShowModal(true);
  };

  const deleteItem = async (id: number): Promise<void> => {
    if (confirm('Are you sure you want to delete this item?')) {
      try {
        setLoading(true);
        setError(null);
        
        const success = await raidService.deleteRaidItem(id);
        
        if (success) {
          setSuccess('Item deleted successfully');
          await loadRaidItems(); // Reload data
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
      
      if (editingId) {
        // Update existing item
        const updatedItem = await raidService.updateRaidItem(editingId, item);
        
        if (updatedItem) {
          setSuccess(SUCCESS_MESSAGES.ITEM_UPDATED);
          await loadRaidItems(); // Reload data
        } else {
          setError(ERROR_MESSAGES.UPDATE_FAILED);
        }
      } else {
        // Create new item
        const newItem = await raidService.createRaidItem(item);
        
        if (newItem) {
          setSuccess(SUCCESS_MESSAGES.ITEM_CREATED);
          await loadRaidItems(); // Reload data
        } else {
          setError(ERROR_MESSAGES.CREATE_FAILED);
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
            <h2>Select RAID Type</h2>
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
                <p>Potential threats</p>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Opportunity')}>
                <h3>Opportunity</h3>
                <p>Potential benefits</p>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Issue')}>
                <h3>Issue</h3>
                <p>Current problems</p>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Assumption')}>
                <h3>Assumption</h3>
                <p>Project assumptions</p>
              </div>
              <div className={styles.typeCard} onClick={() => selectType('Dependency')}>
                <h3>Dependency</h3>
                <p>Project dependencies</p>
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
          <PivotItem headerText="All" itemKey="all" />
          <PivotItem headerText="Risk" itemKey="Risk" />
          <PivotItem headerText="Opportunity" itemKey="Opportunity" />
          <PivotItem headerText="Issue" itemKey="Issue" />
          <PivotItem headerText="Assumption" itemKey="Assumption" />
          <PivotItem headerText="Dependency" itemKey="Dependency" />
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