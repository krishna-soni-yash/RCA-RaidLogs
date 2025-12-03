import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Spinner,
  Stack,
  IStackTokens,
  IconButton,
  PrimaryButton
} from '@fluentui/react';

import styles from '../LlBpRc.module.scss';
import { IReusableComponents } from '../../../../../models/Ll Bp Rc/ReusableComponents';
import {
  addReusableComponents,
  fetchReusableComponents,
  updateReusableComponents
} from '../../../../../repositories/LlBpRcrepository';
import ReusableComponentsForm from './ReusableComponentsForm';

interface IReusableComponentsProps {
  context: WebPartContext;
}

const stackTokens: IStackTokens = { childrenGap: 12 };

const ReusableComponents: React.FC<IReusableComponentsProps> = ({ context }) => {
  const [items, setItems] = React.useState<IReusableComponents[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [showForm, setShowForm] = React.useState<boolean>(false);
  const [selectedComponent, setSelectedComponent] = React.useState<IReusableComponents | null>(null);
  const [formMode, setFormMode] = React.useState<'view' | 'edit' | 'create'>('view');
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [formError, setFormError] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const isCreateMode = formMode === 'create';
  const successTimeoutRef = React.useRef<number | undefined>(undefined);

  const clearSuccessMessage = React.useCallback(() => {
    if (successTimeoutRef.current) {
      window.clearTimeout(successTimeoutRef.current);
      successTimeoutRef.current = undefined;
    }
    setSuccessMessage(null);
  }, []);

  const columns: IColumn[] = React.useMemo(() => [
    {
      key: 'actions',
      name: '',
      fieldName: 'actions',
      minWidth: 80,
      maxWidth: 110,
      isResizable: false
    },
    {
      key: 'componentName',
      name: 'Component Name',
      fieldName: 'RcComponentName',
      minWidth: 140,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'location',
      name: 'Location',
      fieldName: 'RcLocation',
      minWidth: 140,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'purpose',
      name: 'Purpose / Functionality',
      fieldName: 'RcPurposeMainFunctionality',
      minWidth: 160,
      maxWidth: 260,
      isResizable: true
    },
    {
      key: 'responsibility',
      name: 'Responsibility',
      fieldName: 'RcResponsibility',
      minWidth: 140,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'remarks',
      name: 'Remarks',
      fieldName: 'RcRemarks',
      minWidth: 140,
      maxWidth: 260,
      isResizable: true
    }
  ], []);

  const handleCloseForm = React.useCallback(() => {
    setShowForm(false);
    setSelectedComponent(null);
    setFormError(null);
    setIsSaving(false);
    setFormMode('view');
  }, []);

  const handleViewItem = React.useCallback((item: IReusableComponents) => {
    setSelectedComponent(item);
    setFormMode('view');
    setFormError(null);
    setShowForm(true);
  }, []);

  const handleCreateClick = React.useCallback(() => {
    setSelectedComponent(null);
    setFormMode('create');
    setFormError(null);
    setShowForm(true);
  }, []);

  const handleEditItem = React.useCallback((item: IReusableComponents) => {
    setSelectedComponent(item);
    setFormMode('edit');
    setFormError(null);
    setShowForm(true);
  }, []);

  const onRenderItemColumn = React.useCallback((item: IReusableComponents, _: number | undefined, column?: IColumn) => {
    if (!column) {
      return null;
    }

    if (column.key === 'actions') {
      const onView = (ev?: any) => {
        ev?.stopPropagation();
        handleViewItem(item);
      };

      const onEdit = (ev?: any) => {
        ev?.stopPropagation();
        handleEditItem(item);
      };

      return (
        <div>
          <IconButton iconProps={{ iconName: 'View' }} ariaLabel="View" onClick={onView} />
          <IconButton iconProps={{ iconName: 'Edit' }} ariaLabel="Edit" onClick={onEdit} />
        </div>
      );
    }

    const fieldName = column.fieldName as keyof IReusableComponents;
    const rawValue = item[fieldName];
    if (rawValue === undefined || rawValue === null || rawValue === '') {
      return <span>-</span>;
    }

    return <span>{typeof rawValue === 'string' ? rawValue : String(rawValue)}</span>;
  }, [handleEditItem, handleViewItem]);

  React.useEffect(() => {
    let isDisposed = false;

    const loadComponents = async (): Promise<void> => {
      setIsLoading(true);
      setError(null);

      try {
        const data = await fetchReusableComponents(true, context);
        if (isDisposed) {
          return;
        }

        const sorted = [...(data ?? [])].sort((a, b) => (b.ID ?? 0) - (a.ID ?? 0));
        setItems(sorted);
      } catch (err: unknown) {
        if (!isDisposed) {
          const message = err instanceof Error ? err.message : typeof err === 'string' ? err : 'Unable to load reusable components.';
          setError(message);
          setItems([]);
        }
      } finally {
        if (!isDisposed) {
          setIsLoading(false);
        }
      }
    };

    void loadComponents();

    return () => {
      isDisposed = true;
    };
  }, [context]);

  React.useEffect(() => {
    return () => {
      if (successTimeoutRef.current) {
        window.clearTimeout(successTimeoutRef.current);
        successTimeoutRef.current = undefined;
      }
    };
  }, []);

  const handleCreateComponent = React.useCallback(async (values: IReusableComponents) => {
    try {
      setIsSaving(true);
      setFormError(null);
      const saved = await addReusableComponents(values, context);
      setItems(prev => {
        const next = [saved, ...prev];
        return next.sort((a, b) => (b.ID ?? 0) - (a.ID ?? 0));
      });
      clearSuccessMessage();
      setSuccessMessage('Save Complete');
      successTimeoutRef.current = window.setTimeout(() => {
        clearSuccessMessage();
      }, 1000);
      handleCloseForm();
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : typeof err === 'string' ? err : 'Failed to save reusable component.';
      setFormError(message);
    } finally {
      setIsSaving(false);
    }
  }, [clearSuccessMessage, context, handleCloseForm]);

  const handleUpdateComponent = React.useCallback(async (values: IReusableComponents) => {
    if (!selectedComponent?.ID) {
      setFormError('Unable to determine which reusable component to update.');
      return;
    }

    try {
      setIsSaving(true);
      setFormError(null);
      const updated = await updateReusableComponents({ ...values, ID: selectedComponent.ID }, context);
      setItems(prev => {
        const next = prev.map(it => (it.ID === updated.ID ? { ...it, ...updated } : it));
        return next.sort((a, b) => (b.ID ?? 0) - (a.ID ?? 0));
      });
      clearSuccessMessage();
      setSuccessMessage('Save Complete');
      successTimeoutRef.current = window.setTimeout(() => {
        clearSuccessMessage();
      }, 1000);
      handleCloseForm();
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : typeof err === 'string' ? err : 'Failed to update reusable component.';
      setFormError(message);
    } finally {
      setIsSaving(false);
    }
  }, [clearSuccessMessage, context, handleCloseForm, selectedComponent]);

  return (
    <div>
      <PrimaryButton
        text="Add Reusable Component"
        onClick={handleCreateClick}
        style={{ marginTop: '8px' }}
      />
      <Stack tokens={stackTokens} className={styles.formWrapper}>
        {successMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={clearSuccessMessage}
          >
            {successMessage}
          </MessageBar>
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {error}
          </MessageBar>
        )}

        {isLoading && <Spinner label="Loading reusable components..." />}

        {!isLoading && !error && (
          items.length > 0 ? (
            <DetailsList
              items={items}
              columns={columns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              onRenderItemColumn={onRenderItemColumn}
              compact
            />
          ) : (
            <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
              No reusable components have been captured yet.
            </MessageBar>
          )
        )}

        {showForm && (
          <div className={styles.overlay} onClick={handleCloseForm}>
            <div className={styles.container} onClick={(e) => e.stopPropagation()}>
              <div className={styles.header}>
                <h3 className={styles.title}>
                  {formMode === 'view' ? 'View Reusable Component' : 'Reusable Component'}
                </h3>
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  ariaLabel="Close"
                  onClick={handleCloseForm}
                  className={styles.closeButton}
                />
              </div>
              {formError && (
                <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                  {formError}
                </MessageBar>
              )}
              <div className={styles.formWrapper}>
                <ReusableComponentsForm
                  mode={formMode}
                  initialValues={selectedComponent ?? undefined}
                  onSubmit={isCreateMode ? handleCreateComponent : formMode === 'edit' ? handleUpdateComponent : undefined}
                  onCancel={handleCloseForm}
                  isSaving={isSaving}
                />
              </div>
            </div>
          </div>
        )}
      </Stack>
    </div>
  );
};

export default ReusableComponents;