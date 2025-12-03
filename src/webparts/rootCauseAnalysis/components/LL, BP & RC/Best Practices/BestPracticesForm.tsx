import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import {
  DefaultButton,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
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

type PersonSelection = {
  id?: number;
  email?: string;
  loginName?: string;
  displayName?: string;
};

type BestPracticesFormState = {
  BpBestPracticesDescription: string;
  BpReferences: string;
  BpResponsibility: string;
  BpResponsibilityId?: number;
  BpResponsibilityEmail?: string;
  BpResponsibilityLoginName?: string;
  BpRemarks: string;
  DataType: string;
};

type BestPracticesFormErrors = {
  BpBestPracticesDescription: string;
  BpReferences: string;
  BpResponsibility: string;
  BpRemarks: string;
};

const fieldDefaults: BestPracticesFormState = {
  BpBestPracticesDescription: '',
  BpReferences: '',
  BpResponsibility: '',
  BpRemarks: '',
  DataType: BestPracticesDataType
};

const formStackTokens: IStackTokens = { childrenGap: 8 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };

const displayFromSelection = (selection: PersonSelection[]): string =>
  selection
    .map(person => person.displayName || person.email || person.loginName || '')
    .filter(Boolean)
    .join('; ');

const toArray = <T,>(value?: T | T[]): T[] => {
  if (value === undefined || value === null) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
};

const pickFirstString = (candidates: any[]): string | undefined => {
  for (const candidate of candidates) {
    if (typeof candidate === 'string') {
      const trimmed = candidate.trim();
      if (trimmed.length > 0) {
        return trimmed;
      }
    }
  }
  return undefined;
};

const mapPickerItemsToSelection = (items: any[]): PersonSelection[] => {
  if (!items || !Array.isArray(items)) {
    return [];
  }

  return items
    .map<PersonSelection | null>((item: any) => {
      if (!item) {
        return null;
      }

      const rawId = item.id ?? item.Id ?? item.ID;
      const numericId = typeof rawId === 'number'
        ? rawId
        : (typeof rawId === 'string' && rawId.trim().length ? Number(rawId) : undefined);

      const email = pickFirstString([item.secondaryText, item.mail, item.Email, item.EMail]);
      const loginName = pickFirstString([item.loginName, item.LoginName, item.UserName]);
      const displayName = pickFirstString([item.text, item.primaryText, item.DisplayName, item.Title, item.name, item.Name]);

      if (numericId === undefined && !email && !loginName && !displayName) {
        return null;
      }

      return {
        id: numericId !== undefined && !isNaN(numericId) ? numericId : undefined,
        email,
        loginName,
        displayName
      };
    })
    .filter((person): person is PersonSelection => person !== null);
};

const BestPracticesForm: React.FC<IBestPracticesFormProps> = ({
  initialValues,
  onSubmit,
  onCancel,
  isSaving,
  mode = 'create',
  context
}) => {
  const [formState, setFormState] = useState<BestPracticesFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<BestPracticesFormErrors>({
    BpBestPracticesDescription: '',
    BpReferences: '',
    BpResponsibility: '',
    BpRemarks: ''
  });
  const [selectedResponsibility, setSelectedResponsibility] = useState<PersonSelection[]>([]);
  const [peoplePickerKey, setPeoplePickerKey] = useState<number>(0);
  const isReadOnly = mode === 'view';

  const mapInitialResponsibility = useCallback((): PersonSelection[] => {
    const ids = toArray<any>(initialValues?.BpResponsibilityId);
    const emails = toArray<any>(initialValues?.BpResponsibilityEmail);
    const logins = toArray<any>(initialValues?.BpResponsibilityLoginName);
    const names = toArray<any>(initialValues?.BpResponsibility);

    const maxLen = Math.max(ids.length, emails.length, logins.length, names.length);
    if (maxLen === 0) {
      if (typeof initialValues?.BpResponsibility === 'string' && initialValues.BpResponsibility.trim().length) {
        return [{ displayName: initialValues.BpResponsibility.trim() }];
      }
      return [];
    }

    const selections: PersonSelection[] = [];
    for (let index = 0; index < maxLen; index += 1) {
      const rawId = ids[index];
      const rawEmail = emails[index];
      const rawLogin = logins[index];
      const rawName = names[index];

      const numericId = typeof rawId === 'number'
        ? rawId
        : (typeof rawId === 'string' && rawId.trim().length ? Number(rawId) : undefined);

      const email = typeof rawEmail === 'string' && rawEmail.trim().length ? rawEmail.trim() : undefined;
      const loginName = typeof rawLogin === 'string' && rawLogin.trim().length ? rawLogin.trim() : undefined;
      const displayName = typeof rawName === 'string' && rawName.trim().length ? rawName.trim() : undefined;

      if (numericId || email || loginName || displayName) {
        selections.push({
          id: numericId !== undefined && !isNaN(numericId) ? numericId : undefined,
          email,
          loginName,
          displayName
        });
      }
    }

    return selections;
  }, [initialValues?.BpResponsibility, initialValues?.BpResponsibilityEmail, initialValues?.BpResponsibilityId, initialValues?.BpResponsibilityLoginName]);

  const createInitialState = useCallback((): BestPracticesFormState => {
    const responsibility = mapInitialResponsibility();
    const primary = responsibility[0];

    return {
      BpBestPracticesDescription: initialValues?.BpBestPracticesDescription ?? '',
      BpReferences: initialValues?.BpReferences ?? '',
      BpResponsibility: displayFromSelection(responsibility),
      BpResponsibilityId: primary?.id ?? (typeof initialValues?.BpResponsibilityId === 'number' ? initialValues?.BpResponsibilityId : undefined),
      BpResponsibilityEmail: primary?.email ?? (typeof initialValues?.BpResponsibilityEmail === 'string' ? initialValues?.BpResponsibilityEmail : undefined),
      BpResponsibilityLoginName: primary?.loginName ?? (typeof initialValues?.BpResponsibilityLoginName === 'string' ? initialValues?.BpResponsibilityLoginName : initialValues?.BpResponsibilityEmail as string | undefined),
      BpRemarks: initialValues?.BpRemarks ?? '',
      DataType: initialValues?.DataType ?? BestPracticesDataType
    };
  }, [initialValues, mapInitialResponsibility]);

  const resetState = useCallback(() => {
    const nextState = createInitialState();
    setFormState(nextState);
    setErrors({
      BpBestPracticesDescription: '',
      BpReferences: '',
      BpResponsibility: '',
      BpRemarks: ''
    });

    const responsibility = mapInitialResponsibility();
    setSelectedResponsibility(responsibility);
    setPeoplePickerKey(prev => prev + 1);
  }, [createInitialState, mapInitialResponsibility]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  const validate = useCallback((state: BestPracticesFormState, responsibility: PersonSelection[]) => {
    const hasResponsibility = responsibility.length > 0 && (
      responsibility[0].id !== undefined ||
      (responsibility[0].email !== undefined && responsibility[0].email !== '') ||
      (responsibility[0].loginName !== undefined && responsibility[0].loginName !== '') ||
      (responsibility[0].displayName !== undefined && responsibility[0].displayName !== '')
    );

    const nextErrors: BestPracticesFormErrors = {
      BpBestPracticesDescription: state.BpBestPracticesDescription.trim() ? '' : 'Description is required.',
      BpReferences: '',
      BpResponsibility: hasResponsibility ? '' : 'Responsibility is required.',
      BpRemarks: ''
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
    (field: 'BpBestPracticesDescription' | 'BpReferences' | 'BpResponsibility' | 'BpRemarks') => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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

  const handleResponsibilityChange = useCallback((items: any[]) => {
    if (isReadOnly) {
      return;
    }

    const mapped = mapPickerItemsToSelection(items);
    setSelectedResponsibility(mapped);

    const primary = mapped[0];
    const display = displayFromSelection(mapped);

    setFormState(prev => ({
      ...prev,
      BpResponsibility: display,
      BpResponsibilityId: primary?.id,
      BpResponsibilityEmail: primary?.email,
      BpResponsibilityLoginName: primary?.loginName ?? primary?.email
    }));

    if (mapped.length > 0 && errors.BpResponsibility) {
      setErrors(prev => ({ ...prev, BpResponsibility: '' }));
    }
  }, [errors.BpResponsibility, isReadOnly]);

  const handleSubmit = useCallback(
    (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      if (isReadOnly) {
        onCancel?.();
        return;
      }
      const nextState = { ...formState };
      const responsibilitySnapshot = [...selectedResponsibility];
      if (!validate(nextState, responsibilitySnapshot)) {
        return;
      }

      const primary = responsibilitySnapshot[0];
      const responsibilityId = primary?.id;
      const responsibilityEmail = primary?.email;
      const responsibilityLogin = primary?.loginName ?? primary?.email;

      onSubmit?.({
        BpBestPracticesDescription: nextState.BpBestPracticesDescription.trim(),
        BpReferences: nextState.BpReferences.trim(),
        BpResponsibility: nextState.BpResponsibility.trim(),
        BpResponsibilityId: responsibilityId,
        BpResponsibilityEmail: responsibilityEmail,
        BpResponsibilityLoginName: responsibilityLogin,
        BpRemarks: nextState.BpRemarks.trim(),
        DataType: BestPracticesDataType
      });
    },
    [formState, isReadOnly, onCancel, onSubmit, selectedResponsibility, validate]
  );

  const handleReset = useCallback(() => {
    if (isReadOnly) {
      return;
    }
    resetState();
  }, [isReadOnly, resetState]);

  const defaultSelectedUsers = selectedResponsibility
    .map(person => person.email || person.loginName || person.displayName)
    .filter((value): value is string => typeof value === 'string' && value.length > 0);

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
        {isReadOnly ? (
          <TextField
            label="Responsibility"
            value={formState.BpResponsibility}
            readOnly
          />
        ) : (
          <div>
            <label style={{ display: 'block', marginBottom: 4, fontWeight: 600 }}>Responsibility<span style={{ color: '#a4262c' }}> *</span></label>
            <PeoplePicker
              key={`bp-responsibility-${peoplePickerKey}`}
              context={context as any}
              titleText=""
              defaultSelectedUsers={defaultSelectedUsers}
              showtooltip
              personSelectionLimit={1}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              ensureUser
              onChange={handleResponsibilityChange}
            />
            {errors.BpResponsibility && (
              <span style={{ color: '#a4262c', fontSize: 12, marginTop: 4, display: 'block' }}>{errors.BpResponsibility}</span>
            )}
          </div>
        )}
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