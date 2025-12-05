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
import { ReusableComponentsDataType, type IReusableComponents } from '../../../../../models/Ll Bp Rc/ReusableComponents';

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
  const isReadOnly = mode === 'view';

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
  }, [createInitialState]);

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
        RcComponentName: nextState.RcComponentName.trim(),
        RcLocation: nextState.RcLocation.trim(),
        RcPurposeMainFunctionality: nextState.RcPurposeMainFunctionality.trim(),
        RcRemarks: nextState.RcRemarks.trim(),
        DataType: ReusableComponentsDataType
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

export default ReusableComponentsForm;