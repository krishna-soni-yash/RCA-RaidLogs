import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
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
import { IReusableComponents, ReusableComponentsDataType } from '../../../../../models/Ll Bp Rc/ReusableComponents';

export interface IReusableComponentsFormProps {
  initialValues?: Partial<IReusableComponents>;
  onSubmit?: (values: IReusableComponents) => void;
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

type ReusableComponentsFormState = {
  RcComponentName: string;
  RcLocation: string;
  RcPurposeMainFunctionality: string;
  RcResponsibility: string;
  RcResponsibilityId?: number;
  RcResponsibilityEmail?: string;
  RcResponsibilityLoginName?: string;
  RcRemarks: string;
  DataType: string;
};

type ReusableComponentsFormErrors = {
  RcComponentName: string;
  RcLocation: string;
  RcPurposeMainFunctionality: string;
  RcResponsibility: string;
  RcRemarks: string;
};

class PeoplePickerErrorBoundary extends React.Component<{ onError?: () => void; children: React.ReactNode }, { hasError: boolean }> {
  constructor(props: { onError?: () => void; children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError() {
    return { hasError: true };
  }

  componentDidCatch(error: any, info: any) {
    console.error('PeoplePicker render error caught by boundary', error, info);
    this.props.onError?.();
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      return null;
    }
    return this.props.children;
  }
}

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

const ReusableComponentsForm: React.FC<IReusableComponentsFormProps> = ({
  initialValues,
  onSubmit,
  onCancel,
  isSaving,
  mode = 'create',
  context
}) => {
  const [formState, setFormState] = useState<ReusableComponentsFormState>({ ...fieldDefaults });
  const [errors, setErrors] = useState<ReusableComponentsFormErrors>({
    RcComponentName: '',
    RcLocation: '',
    RcPurposeMainFunctionality: '',
    RcResponsibility: '',
    RcRemarks: ''
  });
  const [selectedResponsibility, setSelectedResponsibility] = useState<PersonSelection[]>([]);
  const [peoplePickerKey, setPeoplePickerKey] = useState<number>(0);
  const [peoplePickerFailed, setPeoplePickerFailed] = useState<boolean>(false);
  const isReadOnly = mode === 'view';
  const peoplePickerWebUrl = useMemo(() => {
    return (context as any)?.pageContext?.web?.absoluteUrl ?? window.location.origin;
  }, [context]);

  const mapInitialResponsibility = useCallback((): PersonSelection[] => {
    const ids = toArray<any>(initialValues?.RcResponsibilityId);
    const emails = toArray<any>(initialValues?.RcResponsibilityEmail);
    const logins = toArray<any>(initialValues?.RcResponsibilityLoginName);
    const names = toArray<any>(initialValues?.RcResponsibility);

    const maxLen = Math.max(ids.length, emails.length, logins.length, names.length);
    if (maxLen === 0) {
      if (typeof initialValues?.RcResponsibility === 'string' && initialValues.RcResponsibility.trim().length) {
        return [{ displayName: initialValues.RcResponsibility.trim() }];
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
  }, [initialValues?.RcResponsibility, initialValues?.RcResponsibilityEmail, initialValues?.RcResponsibilityId, initialValues?.RcResponsibilityLoginName]);

  const createInitialState = useCallback((): ReusableComponentsFormState => {
    const responsibility = mapInitialResponsibility();
    const primary = responsibility[0];

    return {
      RcComponentName: initialValues?.RcComponentName ?? '',
      RcLocation: initialValues?.RcLocation ?? '',
      RcPurposeMainFunctionality: initialValues?.RcPurposeMainFunctionality ?? '',
      RcResponsibility: displayFromSelection(responsibility),
      RcResponsibilityId: primary?.id ?? (typeof initialValues?.RcResponsibilityId === 'number' ? initialValues?.RcResponsibilityId : undefined),
      RcResponsibilityEmail: primary?.email ?? (typeof initialValues?.RcResponsibilityEmail === 'string' ? initialValues?.RcResponsibilityEmail : undefined),
      RcResponsibilityLoginName: primary?.loginName ?? (typeof initialValues?.RcResponsibilityLoginName === 'string' ? initialValues?.RcResponsibilityLoginName : initialValues?.RcResponsibilityEmail as string | undefined),
      RcRemarks: initialValues?.RcRemarks ?? '',
      DataType: initialValues?.DataType ?? ReusableComponentsDataType
    };
  }, [initialValues, mapInitialResponsibility]);

  const resetState = useCallback(() => {
    const nextState = createInitialState();
    setFormState(nextState);
    setErrors({
      RcComponentName: '',
      RcLocation: '',
      RcPurposeMainFunctionality: '',
      RcResponsibility: '',
      RcRemarks: ''
    });

    const responsibility = mapInitialResponsibility();
    setSelectedResponsibility(responsibility);
    setPeoplePickerKey(prev => prev + 1);
    setPeoplePickerFailed(false);
  }, [createInitialState, mapInitialResponsibility]);

  useEffect(() => {
    resetState();
  }, [resetState]);

  useEffect(() => {
    const handleWindowError = (ev: ErrorEvent) => {
      const message = ev?.message || ev?.error?.message || '';
      if (message && (message.indexOf('PeopleSearchService') !== -1 || message.indexOf('searchTenant') !== -1)) {
        setPeoplePickerFailed(true);
      }
    };

    const handleUnhandledRejection = (ev: PromiseRejectionEvent) => {
      const reason = ev?.reason;
      const message = typeof reason === 'string' ? reason : (reason?.message || '');
      if (message && (message.indexOf('PeopleSearchService') !== -1 || message.indexOf('searchTenant') !== -1)) {
        setPeoplePickerFailed(true);
      }
    };

    window.addEventListener('error', handleWindowError);
    window.addEventListener('unhandledrejection', handleUnhandledRejection);
    return () => {
      window.removeEventListener('error', handleWindowError);
      window.removeEventListener('unhandledrejection', handleUnhandledRejection);
    };
  }, []);

  const validate = useCallback((state: ReusableComponentsFormState, responsibility: PersonSelection[]) => {
    const hasResponsibility = responsibility.length > 0 && (
      responsibility[0].id !== undefined ||
      (responsibility[0].email !== undefined && responsibility[0].email !== '') ||
      (responsibility[0].loginName !== undefined && responsibility[0].loginName !== '') ||
      (responsibility[0].displayName !== undefined && responsibility[0].displayName !== '')
    );

    const nextErrors: ReusableComponentsFormErrors = {
      RcComponentName: state.RcComponentName.trim() ? '' : 'Component name is required.',
      RcLocation: state.RcLocation.trim() ? '' : 'Location is required.',
      RcPurposeMainFunctionality: state.RcPurposeMainFunctionality.trim() ? '' : 'Purpose or functionality is required.',
      RcResponsibility: hasResponsibility ? '' : 'Responsibility is required.',
      RcRemarks: ''
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
    (field: 'RcComponentName' | 'RcLocation' | 'RcPurposeMainFunctionality' | 'RcResponsibility' | 'RcRemarks') => (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
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
    setPeoplePickerFailed(false);

    const primary = mapped[0];
    const display = displayFromSelection(mapped);

    setFormState(prev => ({
      ...prev,
      RcResponsibility: display,
      RcResponsibilityId: primary?.id,
      RcResponsibilityEmail: primary?.email,
      RcResponsibilityLoginName: primary?.loginName ?? primary?.email
    }));

    if (mapped.length > 0 && errors.RcResponsibility) {
      setErrors(prev => ({ ...prev, RcResponsibility: '' }));
    }
  }, [errors.RcResponsibility, isReadOnly]);

  const handleManualResponsibilityInput = useCallback((_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
    if (isReadOnly) {
      return;
    }

    const trimmed = (value ?? '').trim();
    const containsAtSymbol = trimmed.indexOf('@') !== -1;
    setFormState(prev => ({
      ...prev,
      RcResponsibility: trimmed,
      RcResponsibilityId: undefined,
      RcResponsibilityEmail: containsAtSymbol ? trimmed : undefined,
      RcResponsibilityLoginName: trimmed || undefined
    }));

    if (trimmed) {
      setSelectedResponsibility([
        {
          displayName: trimmed,
          email: containsAtSymbol ? trimmed : undefined,
          loginName: trimmed || undefined
        }
      ]);
      if (errors.RcResponsibility) {
        setErrors(prev => ({ ...prev, RcResponsibility: '' }));
      }
    } else {
      setSelectedResponsibility([]);
    }
  }, [errors.RcResponsibility, isReadOnly]);

  const handleSubmit = useCallback((event: React.FormEvent<HTMLFormElement>) => {
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
      RcComponentName: nextState.RcComponentName.trim(),
      RcLocation: nextState.RcLocation.trim(),
      RcPurposeMainFunctionality: nextState.RcPurposeMainFunctionality.trim(),
      RcResponsibility: nextState.RcResponsibility.trim(),
      RcResponsibilityId: responsibilityId,
      RcResponsibilityEmail: responsibilityEmail,
      RcResponsibilityLoginName: responsibilityLogin,
      RcRemarks: nextState.RcRemarks.trim(),
      DataType: ReusableComponentsDataType
    });
  }, [formState, isReadOnly, onCancel, onSubmit, selectedResponsibility, validate]);

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
        {isReadOnly ? (
          <TextField
            label="Responsibility"
            value={formState.RcResponsibility}
            readOnly
          />
        ) : (
          <div>
            <label style={{ display: 'block', marginBottom: 4, fontWeight: 600 }}>Responsibility<span style={{ color: '#a4262c' }}> *</span></label>
            {!peoplePickerFailed && context ? (
              <PeoplePickerErrorBoundary onError={() => setPeoplePickerFailed(true)}>
                <PeoplePicker
                  key={`rc-responsibility-${peoplePickerKey}`}
                  context={context as any}
                  webAbsoluteUrl={peoplePickerWebUrl}
                  titleText=""
                  defaultSelectedUsers={defaultSelectedUsers}
                  showtooltip
                  personSelectionLimit={1}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={300}
                  ensureUser
                  placeholder="Type a name or email..."
                  onChange={handleResponsibilityChange}
                />
              </PeoplePickerErrorBoundary>
            ) : (
              <>
                <TextField
                  value={formState.RcResponsibility}
                  onChange={handleManualResponsibilityInput}
                  placeholder="Enter name or email manually"
                />
                <span style={{ color: '#605e5c', fontSize: 12, marginTop: 4, display: 'block' }}>
                  People search is unavailable. Enter the responsible person manually.
                </span>
                <DefaultButton
                  style={{ marginTop: 8 }}
                  text="Retry people search"
                  onClick={() => {
                    setPeoplePickerFailed(false);
                    setPeoplePickerKey(prev => prev + 1);
                  }}
                />
              </>
            )}
            {errors.RcResponsibility && (
              <span style={{ color: '#a4262c', fontSize: 12, marginTop: 4, display: 'block' }}>{errors.RcResponsibility}</span>
            )}
          </div>
        )}
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