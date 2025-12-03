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
  IconButton
} from '@fluentui/react';
import styles from '../LlBpRc.module.scss';
import { ILessonsLearnt } from '../../../../../models/Ll Bp Rc/LessonsLearnt';
import { fetchLessonsLearnt } from '../../../../../repositories/LlBpRcrepository';
import LessonsLearntForm from './LessonsLearntForm';

interface ILessonsLearntProps {
  context: WebPartContext;
}

const stackTokens: IStackTokens = { childrenGap: 12 };

const LessonsLearnt: React.FC<ILessonsLearntProps> = ({ context }) => {
  const [items, setItems] = React.useState<ILessonsLearnt[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [showLessonsLearntForm, setShowLessonsLearntForm] = React.useState<boolean>(false);
  const [selectedLesson, setSelectedLesson] = React.useState<ILessonsLearnt | null>(null);
  const [formMode, setFormMode] = React.useState<'view' | 'edit' | 'create'>('view');

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
      key: 'problem',
      name: 'Problem Faced / Learning',
      fieldName: 'LlProblemFacedLearning',
      minWidth: 140,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'category',
      name: 'Category',
      fieldName: 'LlCategory',
      minWidth: 140,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'solution',
      name: 'Solution',
      fieldName: 'LlSolution',
      minWidth: 140,
      maxWidth: 260,
      isResizable: true
    },
    {
      key: 'remarks',
      name: 'Remarks',
      fieldName: 'LlRemarks',
      minWidth: 180,
      maxWidth: 320,
      isResizable: true
    }
  ], []);

  const handleCloseForm = React.useCallback(() => {
    setShowLessonsLearntForm(false);
    setSelectedLesson(null);
  }, []);

  const handleViewItem = React.useCallback((lesson: ILessonsLearnt) => {
    setSelectedLesson(lesson);
    setFormMode('view');
    setShowLessonsLearntForm(true);
  }, []);

  const onRenderItemColumn = React.useCallback((item: ILessonsLearnt, _: number | undefined, column?: IColumn) => {
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
        alert(`Edit Lessons Learnt item ${item.ID}`);
      };

      return (
        <div>
          <IconButton iconProps={{ iconName: 'View' }} ariaLabel="View" onClick={onView} />
          <IconButton iconProps={{ iconName: 'Edit' }} ariaLabel="Edit" onClick={onEdit} />
        </div>
      );
    }

    const fieldName = column.fieldName as keyof ILessonsLearnt;
    const rawValue = item[fieldName];
    if (rawValue === undefined || rawValue === null || rawValue === '') {
      return <span>-</span>;
    }

    return <span>{typeof rawValue === 'string' ? rawValue : String(rawValue)}</span>;
  }, [handleViewItem]);

  React.useEffect(() => {
    let isDisposed = false;

    const loadLessons = async (): Promise<void> => {
      setIsLoading(true);
      setError(null);

      try {
        const data = await fetchLessonsLearnt(true, context);
        if (isDisposed) {
          return;
        }

        const sorted = [...(data ?? [])].sort((a, b) => (b.ID ?? 0) - (a.ID ?? 0));
        setItems(sorted);
      } catch (err: unknown) {
        if (!isDisposed) {
          const message =
            err instanceof Error ? err.message : typeof err === 'string' ? err : 'Unable to load Lessons Learnt.';
          setError(message);
          setItems([]);
        }
      } finally {
        if (!isDisposed) {
          setIsLoading(false);
        }
      }
    };

    // Load data once component is ready.
    void loadLessons();

    return () => {
      isDisposed = true;
    };
  }, [context]);

  return (
    <Stack tokens={stackTokens} className={styles.formWrapper}>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      )}

      {isLoading && <Spinner label="Loading lessons learnt..." />}

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
            No lessons learnt have been captured yet.
          </MessageBar>
        )
      )}
      {showLessonsLearntForm && (
        <div className={styles.overlay} onClick={handleCloseForm}>
          <div className={styles.container} onClick={(e) => e.stopPropagation()}>
            <div className={styles.header}>
              <h3 className={styles.title}>
                {formMode === 'view' ? 'View Lessons Learnt' : 'Lessons Learnt Form'}
              </h3>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                ariaLabel="Close"
                onClick={handleCloseForm}
                className={styles.closeButton}
              />
            </div>
            <div className={styles.formWrapper}>
              <LessonsLearntForm
                mode={formMode}
                initialValues={selectedLesson ?? undefined}
                onCancel={handleCloseForm}
              />
            </div>
          </div>
        </div>
      )}
    </Stack>
  );
};

export default LessonsLearnt;