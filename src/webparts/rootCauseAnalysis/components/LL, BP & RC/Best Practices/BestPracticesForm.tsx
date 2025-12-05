/*eslint-disable*/
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  Label,
  Link,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import styles from '../LlBpRc.module.scss';
import { type IBestPracticeAttachment, type IBestPractices, BestPracticesDataType } from '../../../../../models/Ll Bp Rc/BestPractices';

export interface IBestPracticesFormProps {
  initialValues?: Partial<IBestPractices>;
  onSubmit?: (values: IBestPractices) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
  context: WebPartContext;
}

type BestPracticesFormState = {
  BpBestPracticesDescription: string;
  BpCategory: string;
  BpReferences: string;
  BpRemarks: string;
  DataType: string;
};

type BestPracticesFormErrors = {
  BpBestPracticesDescription: string;
  BpCategory: string;
  BpReferences: string;
  BpRemarks: string;
};

const fieldDefaults: BestPracticesFormState = {
  BpBestPracticesDescription: '',
  BpCategory: '',
  BpReferences: '',
  BpRemarks: '',
  DataType: BestPracticesDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };
const attachmentListTokens: IStackTokens = { childrenGap: 4 };
const attachmentRowTokens: IStackTokens = { childrenGap: 8 };

const BestPracticesForm: React.FC<IBestPracticesFormProps> = (props) => {
  const {
    initialValues,
    onSubmit,
    onCancel,
    isSaving,
    mode = 'create'
  } = props;
  const [formState, setFormState] = useState<BestPracticesFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<BestPracticesFormErrors>({
    BpBestPracticesDescription: '',
    BpCategory: '',
    BpReferences: '',
    BpRemarks: ''
  });
  const [existingAttachments, setExistingAttachments] = useState<IBestPracticeAttachment[]>([]);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const isReadOnly = mode === 'view';
  const attachmentInputId = React.useMemo(() => `bp-attachments-${Math.random().toString(36).slice(2)}`, []);

  const createInitialState = useCallback((): BestPracticesFormState => ({
    BpBestPracticesDescription: initialValues?.BpBestPracticesDescription ?? '',
    BpCategory: initialValues?.BpCategory ?? '',
    BpReferences: initialValues?.BpReferences ?? '',
    BpRemarks: initialValues?.BpRemarks ?? '',
    DataType: initialValues?.DataType ?? BestPracticesDataType
  }), [initialValues]);

  const resetState = useCallback(() => {
    const nextState = createInitialState();
    setFormState(nextState);
    setErrors({
      BpBestPracticesDescription: '',
      BpCategory: '',
      BpReferences: '',
      BpRemarks: ''
    });
    setExistingAttachments(Array.isArray(initialValues?.attachments) ? initialValues.attachments.map(att => ({ ...att })) : []);
    setNewAttachments([]);
  }, [createInitialState, initialValues]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: BestPracticesFormState) => {
    const nextErrors: BestPracticesFormErrors = {
      BpBestPracticesDescription: state.BpBestPracticesDescription.trim() ? '' : 'Description is required.',
      BpCategory: state.BpCategory.trim() ? '' : 'Category is required.',
      BpReferences: state.BpReferences.trim() ? '' : 'References are required.',
      BpRemarks: ''
    };

    setErrors(nextErrors);

    return (
      nextErrors.BpBestPracticesDescription === '' &&
      nextErrors.BpCategory === '' &&
      nextErrors.BpReferences === '' &&
      nextErrors.BpRemarks === ''
    );
  }, []);

  const handleChange = useCallback(
    (field: 'BpBestPracticesDescription' | 'BpCategory' | 'BpReferences' | 'BpRemarks') => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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
        BpBestPracticesDescription: nextState.BpBestPracticesDescription.trim(),
        BpCategory: nextState.BpCategory.trim(),
        BpReferences: nextState.BpReferences.trim(),
        BpRemarks: nextState.BpRemarks.trim(),
        DataType: BestPracticesDataType,
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
          label="Best Practice Description"
          value={formState.BpBestPracticesDescription}
          onChange={isReadOnly ? undefined : handleChange('BpBestPracticesDescription')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.BpBestPracticesDescription}
          readOnly={isReadOnly}
        />
        <TextField
          label="Category"
          value={formState.BpCategory}
          onChange={isReadOnly ? undefined : handleChange('BpCategory')}
          required
          errorMessage={errors.BpCategory}
          readOnly={isReadOnly}
        />
        <TextField
          label="References"
          value={formState.BpReferences}
          onChange={isReadOnly ? undefined : handleChange('BpReferences')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.BpReferences}
          readOnly={isReadOnly}
        />
        <TextField
          label="Remarks"
          value={formState.BpRemarks}
          onChange={isReadOnly ? undefined : handleChange('BpRemarks')}
          multiline
          autoAdjustHeight
          errorMessage={errors.BpRemarks}
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

        <Stack horizontal tokens={buttonStackTokens}>
          {!isReadOnly && (
            <>
              <PrimaryButton
                type="submit"
                text={isSaving ? 'Saving…' : 'Save'}
                disabled={!!isSaving}
              />
              <DefaultButton type="button" text="Reset" onClick={handleReset} disabled={isSaving} />
            </>
          )}
        </Stack>
      </Stack>
    </form>
  );
};

export default BestPracticesForm;