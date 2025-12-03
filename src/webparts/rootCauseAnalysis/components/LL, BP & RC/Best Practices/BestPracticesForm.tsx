import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';

import styles from '../LlBpRc.module.scss';
import { IBestPractices, BestPracticesDataType } from '../../../../../models/Ll Bp Rc/BestPractices';

export interface IBestPracticesFormProps {
  initialValues?: Partial<IBestPractices>;
  onSubmit?: (values: IBestPractices) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
}

type BestPracticesFormState = Required<Omit<IBestPractices, 'ID' | 'DataType'>> & { DataType: string };

const fieldDefaults: BestPracticesFormState = {
  BpBestPracticesDescription: '',
  BpReferences: '',
  BpResponsibility: '',
  BpRemarks: '',
  DataType: BestPracticesDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };

const BestPracticesForm: React.FC<IBestPracticesFormProps> = ({
  initialValues,
  onSubmit,
  onCancel,
  isSaving,
  mode = 'create'
}) => {
  const [formState, setFormState] = useState<BestPracticesFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<Record<keyof BestPracticesFormState, string>>({
    BpBestPracticesDescription: '',
    BpReferences: '',
    BpResponsibility: '',
    BpRemarks: '',
    DataType: ''
  });
  const isReadOnly = mode === 'view';

  const createInitialState = useCallback((): BestPracticesFormState => ({
    BpBestPracticesDescription: initialValues?.BpBestPracticesDescription ?? '',
    BpReferences: initialValues?.BpReferences ?? '',
    BpResponsibility: initialValues?.BpResponsibility ?? '',
    BpRemarks: initialValues?.BpRemarks ?? '',
    DataType: initialValues?.DataType ?? BestPracticesDataType
  }), [initialValues]);

  const resetState = useCallback(() => {
    setFormState(createInitialState());
    setErrors({
      BpBestPracticesDescription: '',
      BpReferences: '',
      BpResponsibility: '',
      BpRemarks: '',
      DataType: ''
    });
  }, [createInitialState]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: BestPracticesFormState) => {
    const nextErrors: Record<keyof BestPracticesFormState, string> = {
      BpBestPracticesDescription: state.BpBestPracticesDescription ? '' : 'Description is required.',
      BpReferences: '',
      BpResponsibility: state.BpResponsibility ? '' : 'Responsibility is required.',
      BpRemarks: '',
      DataType: ''
    };
    setErrors(nextErrors);
    return (
      nextErrors.BpBestPracticesDescription === '' &&
      nextErrors.BpReferences === '' &&
      nextErrors.BpResponsibility === '' &&
      nextErrors.BpRemarks === ''
    );
  }, []);

  const handleChange = useCallback(
    (field: keyof BestPracticesFormState) => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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
        BpReferences: nextState.BpReferences.trim(),
        BpResponsibility: nextState.BpResponsibility.trim(),
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
          label="References"
          value={formState.BpReferences}
          onChange={isReadOnly ? undefined : handleChange('BpReferences')}
          multiline
          autoAdjustHeight
          errorMessage={errors.BpReferences}
          readOnly={isReadOnly}
        />
        <TextField
          label="Responsibility"
          value={formState.BpResponsibility}
          onChange={isReadOnly ? undefined : handleChange('BpResponsibility')}
          required
          errorMessage={errors.BpResponsibility}
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
              <PrimaryButton type="submit" text={isSaving ? 'Saving…' : 'Save'} />
              <DefaultButton type="button" text="Reset" onClick={handleReset} disabled={isSaving} />
            </>
          )}
        </Stack>
      </Stack>
    </form>
  );
};

export default BestPracticesForm;