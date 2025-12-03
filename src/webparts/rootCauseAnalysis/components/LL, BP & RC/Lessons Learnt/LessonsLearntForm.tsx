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
import { ILessonsLearnt } from '../../../../../models/Ll Bp Rc/LessonsLearnt';

export interface ILessonsLearntFormProps {
  initialValues?: Partial<ILessonsLearnt>;
  onSubmit?: (values: ILessonsLearnt) => void;
  onCancel?: () => void;
  isSaving?: boolean;
  mode?: 'create' | 'edit' | 'view';
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
  isSaving,
  mode = 'create'
}) => {
  const [formState, setFormState] = useState<LessonsLearntFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<Record<keyof LessonsLearntFormState, string>>({
    LlProblemFacedLearning: '',
    LlCategory: '',
    LlSolution: '',
    LlRemarks: ''
  });
  const isReadOnly = mode === 'view';

  const createInitialState = useCallback((): LessonsLearntFormState => ({
    LlProblemFacedLearning: initialValues?.LlProblemFacedLearning ?? '',
    LlCategory: initialValues?.LlCategory ?? '',
    LlSolution: initialValues?.LlSolution ?? '',
    LlRemarks: initialValues?.LlRemarks ?? ''
  }), [initialValues]);

  const resetState = useCallback(() => {
    setFormState(createInitialState());
    setErrors({
      LlProblemFacedLearning: '',
      LlCategory: '',
      LlSolution: '',
      LlRemarks: ''
    });
  }, [createInitialState]);

  useEffect(() => {
    resetState();
  }, [resetState]);

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
        LlProblemFacedLearning: nextState.LlProblemFacedLearning.trim(),
        LlCategory: nextState.LlCategory.trim(),
        LlSolution: nextState.LlSolution.trim(),
        LlRemarks: nextState.LlRemarks.trim()
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

        <Stack horizontal tokens={buttonStackTokens}>
          {!isReadOnly && (
            <>
              <PrimaryButton type="submit" text={isSaving ? 'Saving…' : 'Save'}/>
              <DefaultButton type="button" text="Reset" onClick={handleReset} disabled={isSaving} />
            </>
          )}
        </Stack>
      </Stack>
    </form>
  );
};

export default LessonsLearntForm;