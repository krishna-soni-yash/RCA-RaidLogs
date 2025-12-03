import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';

import styles from '../LlBpRc.module.scss';
import { ILessonsLearnt } from '../../../../../models/Ll Bp Rc/LessonsLearnt';

export interface ILessonsLearntFormProps {
  initialValues?: Partial<ILessonsLearnt>;
  onSubmit?: (values: ILessonsLearnt) => void;
  onCancel?: () => void;
  isSaving?: boolean;
}

type LessonsLearntFormState = Required<Omit<ILessonsLearnt, 'ID'>>;

const fieldDefaults: LessonsLearntFormState = {
  LlProblemFacedLearning: '',
  LlCategory: '',
  LlSolution: '',
  LlRemarks: ''
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };

const LessonsLearntForm: React.FC<ILessonsLearntFormProps> = ({
  initialValues,
  onSubmit,
  onCancel,
  isSaving
}) => {
  const [formState, setFormState] = useState<LessonsLearntFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<Record<keyof LessonsLearntFormState, string>>({
    LlProblemFacedLearning: '',
    LlCategory: '',
    LlSolution: '',
    LlRemarks: ''
  });

  useEffect(() => {
    if (!initialValues) {
      setFormState({ ...fieldDefaults });
      setErrors({
        LlProblemFacedLearning: '',
        LlCategory: '',
        LlSolution: '',
        LlRemarks: ''
      });
      return;
    }

    setFormState({
      LlProblemFacedLearning: initialValues.LlProblemFacedLearning ?? '',
      LlCategory: initialValues.LlCategory ?? '',
      LlSolution: initialValues.LlSolution ?? '',
      LlRemarks: initialValues.LlRemarks ?? ''
    });
  }, [initialValues]);

  const validate = useCallback((state: LessonsLearntFormState) => {
    const nextErrors: Record<keyof LessonsLearntFormState, string> = {
      LlProblemFacedLearning: state.LlProblemFacedLearning ? '' : 'Please describe the problem or learning.',
      LlCategory: state.LlCategory ? '' : 'Category is required.',
      LlSolution: state.LlSolution ? '' : 'Solution details are required.',
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
    (field: keyof LessonsLearntFormState) => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
      setFormState(prev => ({ ...prev, [field]: value ?? '' }));
      if (errors[field]) {
        setErrors(prev => ({ ...prev, [field]: '' }));
      }
    },
    [errors]
  );

  const handleSubmit = useCallback(
    (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      const nextState = { ...formState };
      if (!validate(nextState)) {
        return;
      }

      onSubmit?.({
        LlProblemFacedLearning: nextState.LlProblemFacedLearning.trim(),
        LlCategory: nextState.LlCategory.trim(),
        LlSolution: nextState.LlSolution.trim(),
        LlRemarks: nextState.LlRemarks.trim()
      });
    },
    [formState, onSubmit, validate]
  );

  const canSubmit = useMemo(() => {
    if (isSaving) {
      return false;
    }

    return (
      formState.LlProblemFacedLearning.trim().length > 0 &&
      formState.LlCategory.trim().length > 0 &&
      formState.LlSolution.trim().length > 0
    );
  }, [formState, isSaving]);

  return (
    <form className={styles.wrapper} onSubmit={handleSubmit} noValidate>
      <Stack tokens={formStackTokens}>
        <TextField
          label="Problem Faced / Learning"
          value={formState.LlProblemFacedLearning}
          onChange={handleChange('LlProblemFacedLearning')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.LlProblemFacedLearning}
        />
        <TextField
          label="Category"
          value={formState.LlCategory}
          onChange={handleChange('LlCategory')}
          required
          errorMessage={errors.LlCategory}
        />
        <TextField
          label="Solution"
          value={formState.LlSolution}
          onChange={handleChange('LlSolution')}
          multiline
          autoAdjustHeight
          required
          errorMessage={errors.LlSolution}
        />
        <TextField
          label="Remarks"
          value={formState.LlRemarks}
          onChange={handleChange('LlRemarks')}
          multiline
          autoAdjustHeight
          errorMessage={errors.LlRemarks}
        />

        <Stack horizontal tokens={buttonStackTokens}>
          <PrimaryButton type="submit" text={isSaving ? 'Saving…' : 'Save'} disabled={!canSubmit} />
          {onCancel && <DefaultButton type="button" text="Cancel" onClick={onCancel} disabled={isSaving} />}
        </Stack>
      </Stack>
    </form>
  );
};

export default LessonsLearntForm;