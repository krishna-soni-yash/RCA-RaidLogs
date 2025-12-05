/*eslint-disable*/
import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import styles from '../LlBpRc.module.scss';
import { IBestPractices, BestPracticesDataType } from '../../../../../models/Ll Bp Rc/BestPractices';

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
  const isReadOnly = mode === 'view';

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
  }, [createInitialState]);

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
        BpBestPracticesDescription: nextState.BpBestPracticesDescription.trim(),
        BpCategory: nextState.BpCategory.trim(),
        BpReferences: nextState.BpReferences.trim(),
        BpRemarks: nextState.BpRemarks.trim(),
        DataType: BestPracticesDataType
      });
    },
    [formState, isReadOnly, onCancel, onSubmit, validate]
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