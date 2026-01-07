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
import { ReusableComponentsDataType, type IReusableComponentAttachment, type IReusableComponents } from '../../../../../models/Ll Bp Rc/ReusableComponents';

export interface IReusableComponentsFormProps {
  initialValues?: Partial<IReusableComponents>;
  onSubmit?: (values: IReusableComponents) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
  context: WebPartContext;
}

type ReusableComponentsFormState = {
  RcComponentName: string;
  RcLocation: string;
  RcPurposeMainFunctionality: string;
  RcRemarks: string;
  DataType: string;
};

type ReusableComponentsFormErrors = {
  RcComponentName: string;
  RcLocation: string;
  RcPurposeMainFunctionality: string;
  RcRemarks: string;
};

const fieldDefaults: ReusableComponentsFormState = {
  RcComponentName: '',
  RcLocation: '',
  RcPurposeMainFunctionality: '',
  RcRemarks: '',
  DataType: ReusableComponentsDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };
const attachmentListTokens: IStackTokens = { childrenGap: 4 };
const attachmentRowTokens: IStackTokens = { childrenGap: 8 };

const ReusableComponentsForm: React.FC<IReusableComponentsFormProps> = (props) => {
  const {
    initialValues,
    onSubmit,
    onCancel,
    isSaving,
    mode = 'create'
  } = props;

  const [formState, setFormState] = useState<ReusableComponentsFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<ReusableComponentsFormErrors>({
    RcComponentName: '',
    RcLocation: '',
    RcPurposeMainFunctionality: '',
    RcRemarks: ''
  });
  const [existingAttachments, setExistingAttachments] = useState<IReusableComponentAttachment[]>([]);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const isReadOnly = mode === 'view';
  const isEditing = mode === 'edit';
  const shouldShowReset = !isReadOnly && !isEditing;
  const attachmentInputId = React.useMemo(() => `rc-attachments-${Math.random().toString(36).slice(2)}`, []);

  const createInitialState = useCallback((): ReusableComponentsFormState => ({
    RcComponentName: initialValues?.RcComponentName ?? '',
    RcLocation: initialValues?.RcLocation ?? '',
    RcPurposeMainFunctionality: initialValues?.RcPurposeMainFunctionality ?? '',
    RcRemarks: initialValues?.RcRemarks ?? '',
    DataType: initialValues?.DataType ?? ReusableComponentsDataType
  }), [initialValues]);

  const resetState = useCallback(() => {
    const nextState = createInitialState();
    setFormState(nextState);
    setErrors({
      RcComponentName: '',
      RcLocation: '',
      RcPurposeMainFunctionality: '',
      RcRemarks: ''
    });
    setExistingAttachments(Array.isArray(initialValues?.attachments) ? initialValues.attachments.map(att => ({ ...att })) : []);
    setNewAttachments([]);
  }, [createInitialState, initialValues]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: ReusableComponentsFormState) => {
    const nextErrors: ReusableComponentsFormErrors = {
      RcComponentName: state.RcComponentName.trim() ? '' : 'Component Name is required.',
      RcLocation: state.RcLocation.trim() ? '' : 'Location is required.',
      RcPurposeMainFunctionality: state.RcPurposeMainFunctionality.trim() ? '' : 'Purpose/Main Functionality is required.',
      RcRemarks: ''
    };

    setErrors(nextErrors);

    return (
      nextErrors.RcComponentName === '' &&
      nextErrors.RcLocation === '' &&
      nextErrors.RcPurposeMainFunctionality === '' &&
      nextErrors.RcRemarks === ''
    );
  }, []);

  const handleChange = useCallback(
    (field: 'RcComponentName' | 'RcLocation' | 'RcPurposeMainFunctionality' | 'RcRemarks') => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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

    // reset input to allow selecting the same file again if needed
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
        RcComponentName: nextState.RcComponentName.trim(),
        RcLocation: nextState.RcLocation.trim(),
        RcPurposeMainFunctionality: nextState.RcPurposeMainFunctionality.trim(),
        RcRemarks: nextState.RcRemarks.trim(),
        DataType: ReusableComponentsDataType,
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
          label="Component Name"
          value={formState.RcComponentName}
          onChange={isReadOnly ? undefined : handleChange('RcComponentName')}
          required
          errorMessage={errors.RcComponentName}
          readOnly={isReadOnly}
        />
        <TextField
          label="Location"
          value={formState.RcLocation}
          onChange={isReadOnly ? undefined : handleChange('RcLocation')}
          required
          errorMessage={errors.RcLocation}
          readOnly={isReadOnly}
        />
        <TextField
          label="Purpose/Main Functionality"
          value={formState.RcPurposeMainFunctionality}
          onChange={isReadOnly ? undefined : handleChange('RcPurposeMainFunctionality')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.RcPurposeMainFunctionality}
          readOnly={isReadOnly}
        />
        <TextField
          label="Remarks"
          value={formState.RcRemarks}
          onChange={isReadOnly ? undefined : handleChange('RcRemarks')}
          multiline
          autoAdjustHeight
          errorMessage={errors.RcRemarks}
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

export default ReusableComponentsForm;