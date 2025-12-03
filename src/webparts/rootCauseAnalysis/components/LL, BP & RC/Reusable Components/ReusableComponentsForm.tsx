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
import { IReusableComponents, ReusableComponentsDataType } from '../../../../../models/Ll Bp Rc/ReusableComponents';

export interface IReusableComponentsFormProps {
  initialValues?: Partial<IReusableComponents>;
  onSubmit?: (values: IReusableComponents) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
}

type ReusableComponentsFormState = Required<Omit<IReusableComponents, 'ID'>>;

const fieldDefaults: ReusableComponentsFormState = {
  RcComponentName: '',
  RcLocation: '',
  RcPurposeMainFunctionality: '',
  RcResponsibility: '',
  RcRemarks: '',
  DataType: ReusableComponentsDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };

const ReusableComponentsForm: React.FC<IReusableComponentsFormProps> = ({
  initialValues,
  onSubmit,
  onCancel,
  isSaving,
  mode = 'create'
}) => {
  const [formState, setFormState] = useState<ReusableComponentsFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<Record<keyof ReusableComponentsFormState, string>>({
    RcComponentName: '',
    RcLocation: '',
    RcPurposeMainFunctionality: '',
    RcResponsibility: '',
    RcRemarks: '',
    DataType: ReusableComponentsDataType
  });
  const isReadOnly = mode === 'view';

  const createInitialState = useCallback((): ReusableComponentsFormState => ({
    RcComponentName: initialValues?.RcComponentName ?? '',
    RcLocation: initialValues?.RcLocation ?? '',
    RcPurposeMainFunctionality: initialValues?.RcPurposeMainFunctionality ?? '',
    RcResponsibility: initialValues?.RcResponsibility ?? '',
    RcRemarks: initialValues?.RcRemarks ?? '',
    DataType: initialValues?.DataType ?? ReusableComponentsDataType
  }), [initialValues]);

  const resetState = useCallback(() => {
    setFormState(createInitialState());
    setErrors({
      RcComponentName: '',
      RcLocation: '',
      RcPurposeMainFunctionality: '',
      RcResponsibility: '',
      RcRemarks: '',
      DataType: ReusableComponentsDataType
    });
  }, [createInitialState]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: ReusableComponentsFormState) => {
    const nextErrors: Record<keyof ReusableComponentsFormState, string> = {
      RcComponentName: state.RcComponentName ? '' : 'Component name is required.',
      RcLocation: state.RcLocation ? '' : 'Location is required.',
      RcPurposeMainFunctionality: state.RcPurposeMainFunctionality ? '' : 'Purpose or functionality is required.',
      RcResponsibility: '',
      RcRemarks: '',
      DataType: ReusableComponentsDataType
    };
    setErrors(nextErrors);
    return (
      nextErrors.RcComponentName === '' &&
      nextErrors.RcLocation === '' &&
      nextErrors.RcPurposeMainFunctionality === '' &&
      nextErrors.RcResponsibility === '' &&
      nextErrors.RcRemarks === ''
    );
  }, []);

  const handleChange = useCallback(
    (field: keyof ReusableComponentsFormState) => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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

  const handleSubmit = useCallback((event: React.FormEvent<HTMLFormElement>) => {
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
      RcComponentName: nextState.RcComponentName.trim(),
      RcLocation: nextState.RcLocation.trim(),
      RcPurposeMainFunctionality: nextState.RcPurposeMainFunctionality.trim(),
      RcResponsibility: nextState.RcResponsibility.trim(),
      RcRemarks: nextState.RcRemarks.trim(),
      DataType: ReusableComponentsDataType
    });
  }, [formState, isReadOnly, onCancel, onSubmit, validate]);

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
          label="Purpose / Main Functionality"
          value={formState.RcPurposeMainFunctionality}
          onChange={isReadOnly ? undefined : handleChange('RcPurposeMainFunctionality')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.RcPurposeMainFunctionality}
          readOnly={isReadOnly}
        />
        <TextField
          label="Responsibility"
          value={formState.RcResponsibility}
          onChange={isReadOnly ? undefined : handleChange('RcResponsibility')}
          errorMessage={errors.RcResponsibility}
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

        <Stack horizontal tokens={buttonStackTokens}>
          {!isReadOnly && (
            <>
              <PrimaryButton type="submit" text={isSaving ? 'Saving…' : 'Save'} />
              <DefaultButton type="button" text="Reset" onClick={handleReset} disabled={!!isSaving} />
            </>
          )}
        </Stack>
      </Stack>
    </form>
  );
};

export default ReusableComponentsForm;