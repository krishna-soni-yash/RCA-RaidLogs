/*eslint-disable*/
import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  Label,
  Link,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';

import styles from '../LlBpRc.module.scss';
import {
  type ILessonsLearnt,
  type ILessonsLearntAttachment,
  LessonsLearntDataType
} from '../../../../../models/Ll Bp Rc/LessonsLearnt';

export interface ILessonsLearntFormProps {
  initialValues?: Partial<ILessonsLearnt>;
  onSubmit?: (values: ILessonsLearnt) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
}

type LessonsLearntFormState = {
  LlProblemFacedLearning: string;
  LlCategory: string;
  LlSolution: string;
  LlRemarks: string;
  DataType: string;
};

type LessonsLearntFormErrors = {
  LlProblemFacedLearning: string;
  LlCategory: string;
  LlSolution: string;
  LlRemarks: string;
};

const fieldDefaults: LessonsLearntFormState = {
  LlProblemFacedLearning: '',
  LlCategory: '',
  LlSolution: '',
  LlRemarks: '',
  DataType: LessonsLearntDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };
const attachmentListTokens: IStackTokens = { childrenGap: 4 };
const attachmentRowTokens: IStackTokens = { childrenGap: 8 };

const LessonsLearntForm: React.FC<ILessonsLearntFormProps> = (props) => {
  const {
    initialValues,
    onSubmit,
    onCancel,
    isSaving,
    mode = 'create'
  } = props;

  const [formState, setFormState] = useState<LessonsLearntFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<LessonsLearntFormErrors>({
    LlProblemFacedLearning: '',
    LlCategory: '',
    LlSolution: '',
    LlRemarks: ''
  });
  const [existingAttachments, setExistingAttachments] = useState<ILessonsLearntAttachment[]>([]);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const isReadOnly = mode === 'view';
  const isEditing = mode === 'edit';
  const shouldShowReset = !isReadOnly && !isEditing;
  const attachmentInputId = useMemo(() => `ll-attachments-${Math.random().toString(36).slice(2)}`, []);

  const createInitialState = useCallback((): LessonsLearntFormState => ({
    LlProblemFacedLearning: initialValues?.LlProblemFacedLearning ?? '',
    LlCategory: initialValues?.LlCategory ?? '',
    LlSolution: initialValues?.LlSolution ?? '',
    LlRemarks: initialValues?.LlRemarks ?? '',
    DataType: initialValues?.DataType ?? LessonsLearntDataType
  }), [initialValues]);

  const resetState = useCallback(() => {
    const nextState = createInitialState();
    setFormState(nextState);
    setErrors({
      LlProblemFacedLearning: '',
      LlCategory: '',
      LlSolution: '',
      LlRemarks: ''
    });
    setExistingAttachments(Array.isArray(initialValues?.attachments) ? initialValues.attachments.map(att => ({ ...att })) : []);
    setNewAttachments([]);
  }, [createInitialState, initialValues]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: LessonsLearntFormState) => {
    const nextErrors: LessonsLearntFormErrors = {
      LlProblemFacedLearning: state.LlProblemFacedLearning.trim() ? '' : 'Please describe the problem or learning.',
      LlCategory: state.LlCategory.trim() ? '' : 'Category is required.',
      LlSolution: state.LlSolution.trim() ? '' : 'Solution details are required.',
      LlRemarks: ''
    };
    setErrors(nextErrors);

    return (
      nextErrors.LlProblemFacedLearning === '' &&
      nextErrors.LlCategory === '' &&
      nextErrors.LlSolution === '' &&
      nextErrors.LlRemarks === ''
    );
  }, []);

  const handleChange = useCallback(
    (field: 'LlProblemFacedLearning' | 'LlCategory' | 'LlSolution' | 'LlRemarks') => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
      if (isReadOnly) {
        return;
      }
      setFormState(prev => ({ ...prev, [field]: value ?? '' }));
      if (errors[field]) {
        setErrors(prev => ({ ...prev, [field]: '' }));
      }
    },
    [errors, isReadOnly]
  );

  const handleAttachmentsAdded = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    if (isReadOnly) {
      return;
    }

    const input = event.currentTarget;
    const files: File[] = [];
    if (input.files && input.files.length > 0) {
      for (let index = 0; index < input.files.length; index++) {
        const file = input.files.item(index);
        if (file) {
          files.push(file);
        }
      }
    }

    if (files.length === 0) {
      return;
    }

    setNewAttachments(prev => {
      const existingNames = new Set<string>([
        ...prev.map(file => file.name),
        ...existingAttachments.map(att => att.FileName)
      ]);

      const next = [...prev];
      for (const file of files) {
        if (!existingNames.has(file.name)) {
          next.push(file);
          existingNames.add(file.name);
        }
      }

      return next;
    });

    input.value = '';
  }, [existingAttachments, isReadOnly]);

  const handleRemoveNewAttachment = useCallback((index: number) => {
    if (isReadOnly) {
      return;
    }

    setNewAttachments(prev => prev.filter((_, idx) => idx !== index));
  }, [isReadOnly]);

  const handleSubmit = useCallback(
    (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      if (isReadOnly) {
        onCancel?.();
        return;
      }

      const nextState = { ...formState };
      if (!validate(nextState)) {
        return;
      }

      onSubmit?.({
        ID: initialValues?.ID,
        LlProblemFacedLearning: nextState.LlProblemFacedLearning.trim(),
        LlCategory: nextState.LlCategory.trim(),
        LlSolution: nextState.LlSolution.trim(),
        LlRemarks: nextState.LlRemarks.trim(),
        DataType: LessonsLearntDataType,
        attachments: existingAttachments.map(att => ({ ...att })),
        newAttachments: [...newAttachments]
      });
    },
    [existingAttachments, formState, initialValues, isReadOnly, newAttachments, onCancel, onSubmit, validate]
  );

  const handleReset = useCallback(() => {
    if (isReadOnly) {
      return;
    }
    resetState();
  }, [isReadOnly, resetState]);

  return (
    <form className={styles.formWrapper} onSubmit={handleSubmit} noValidate>
      <Stack tokens={formStackTokens}>
        <TextField
          label="Problem Faced / Learning"
          value={formState.LlProblemFacedLearning}
          onChange={isReadOnly ? undefined : handleChange('LlProblemFacedLearning')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.LlProblemFacedLearning}
          readOnly={isReadOnly}
        />
        <TextField
          label="Category"
          value={formState.LlCategory}
          onChange={isReadOnly ? undefined : handleChange('LlCategory')}
          required
          errorMessage={errors.LlCategory}
          readOnly={isReadOnly}
        />
        <TextField
          label="Solution"
          value={formState.LlSolution}
          onChange={isReadOnly ? undefined : handleChange('LlSolution')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.LlSolution}
          readOnly={isReadOnly}
        />
        <TextField
          label="Remarks"
          value={formState.LlRemarks}
          onChange={isReadOnly ? undefined : handleChange('LlRemarks')}
          multiline
          autoAdjustHeight
          errorMessage={errors.LlRemarks}
          readOnly={isReadOnly}
        />

        <Stack tokens={attachmentListTokens}>
          <Label htmlFor={attachmentInputId}>Attachments</Label>

          {existingAttachments.length > 0 && (
            <Stack tokens={attachmentListTokens}>
              {existingAttachments.map((attachment, idx) => {
                const key = `existing-${idx}-${attachment.ServerRelativeUrl || attachment.FileName || 'attachment'}`;
                const label = attachment.FileName || attachment.ServerRelativeUrl || 'Attachment';
                if (attachment.ServerRelativeUrl) {
                  return (
                    <Link key={key} href={attachment.ServerRelativeUrl} target="_blank" rel="noreferrer">
                      {label}
                    </Link>
                  );
                }
                return (
                  <span key={key}>{label}</span>
                );
              })}
            </Stack>
          )}

          {newAttachments.length > 0 && (
            <Stack tokens={attachmentListTokens}>
              {newAttachments.map((file, idx) => (
                <Stack key={`new-${idx}-${file.name}`} horizontal tokens={attachmentRowTokens} verticalAlign="center">
                  <span style={{ flexGrow: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.name}</span>
                  {!isReadOnly && (
                    <DefaultButton
                      type="button"
                      text="Remove"
                      onClick={() => handleRemoveNewAttachment(idx)}
                    />
                  )}
                </Stack>
              ))}
            </Stack>
          )}

          {existingAttachments.length === 0 && newAttachments.length === 0 && (
            <span>No attachments added.</span>
          )}

          {!isReadOnly && (
            <input
              id={attachmentInputId}
              type="file"
              multiple
              onChange={handleAttachmentsAdded}
              aria-label="Add attachments"
            />
          )}
        </Stack>

        {!isReadOnly && (
          <Stack horizontal tokens={buttonStackTokens}>
            <PrimaryButton
              type="submit"
              text={isSaving ? 'Saving…' : 'Save'}
              disabled={!!isSaving}
            />
            {shouldShowReset && (
              <DefaultButton type="button" text="Reset" onClick={handleReset} disabled={isSaving} />
            )}
          </Stack>
        )}
      </Stack>
    </form>
  );
};

export default LessonsLearntForm;