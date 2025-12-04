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
  RcResponsibilityId?: number | number[];
  RcResponsibilityEmail?: string | string[];
  RcResponsibilityLoginName?: string | string[];
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

const extractNumberTokens = (value: any): number[] => {
  const results: number[] = [];
  const consume = (input: any): void => {
    if (input === undefined || input === null) {
      return;
    }
    if (Array.isArray(input)) {
      input.forEach(consume);
      return;
    }
    if (typeof input === 'number' && !isNaN(input)) {
      results.push(input);
      return;
    }
    if (typeof input === 'string') {
      const sanitized = input.replace(/;#/g, ';');
      sanitized
        .split(/[;,\n]+/)
        .map(part => part.trim())
        .forEach(part => {
          if (!part) {
            return;
          }
          const parsed = Number(part);
          if (!isNaN(parsed)) {
            results.push(parsed);
          }
        });
      return;
    }
    if (typeof input === 'object') {
      const candidate = (input as any).Id ?? (input as any).ID ?? (input as any).id;
      if (typeof candidate === 'number' && !isNaN(candidate)) {
        results.push(candidate);
      }
      if (Array.isArray((input as any).results)) {
        (input as any).results.forEach(consume);
      }
    }
  };

  consume(value);
  return results;
};

const extractStringTokens = (value: any): string[] => {
  const results: string[] = [];
  const consume = (input: any): void => {
    if (input === undefined || input === null) {
      return;
    }
    if (Array.isArray(input)) {
      input.forEach(consume);
      return;
    }
    if (typeof input === 'string') {
      const sanitized = input.replace(/;#/g, ';');
      sanitized
        .split(/[;,\n]+/)
        .map(part => part.trim())
        .filter(part => part.length > 0 && !/^\d+$/.test(part))
        .forEach(part => results.push(part));
      return;
    }
    if (typeof input === 'object') {
      const candidateStrings = [
        (input as any).displayName,
        (input as any).DisplayName,
        (input as any).Title,
        (input as any).Name,
        (input as any).LoginName,
        (input as any).UserPrincipalName,
        (input as any).Email,
        (input as any).EMail
      ];
      candidateStrings.forEach(candidate => {
        if (typeof candidate === 'string') {
          const trimmed = candidate.trim();
          if (trimmed.length > 0 && !/^\d+$/.test(trimmed)) {
            results.push(trimmed);
          }
        }
      });
      if (Array.isArray((input as any).results)) {
        (input as any).results.forEach(consume);
      }
      return;
    }
    const text = String(input).trim();
    if (text.length > 0 && !/^\d+$/.test(text)) {
      results.push(text);
    }
  };

  consume(value);
  return results;
};

const uniqueNumbers = (values: Array<number | undefined>): number[] => {
  const seen = new Set<number>();
  const unique: number[] = [];
  values.forEach(value => {
    if (typeof value === 'number' && !isNaN(value) && !seen.has(value)) {
      seen.add(value);
      unique.push(value);
    }
  });
  return unique;
};

const uniqueStrings = (values: Array<string | undefined>): string[] => {
  const seen = new Set<string>();
  const unique: string[] = [];
  values.forEach(value => {
    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed.length > 0) {
        const key = trimmed.toLowerCase();
        if (!seen.has(key)) {
          seen.add(key);
          unique.push(trimmed);
        }
      }
    }
  });
  return unique;
};

const collapseNumberValues = (values: Array<number | undefined>): number | number[] | undefined => {
  const unique = uniqueNumbers(values);
  if (unique.length === 0) {
    return undefined;
  }
  return unique.length === 1 ? unique[0] : unique;
};

const collapseStringValues = (values: Array<string | undefined>): string | string[] | undefined => {
  const unique = uniqueStrings(values);
  if (unique.length === 0) {
    return undefined;
  }
  return unique.length === 1 ? unique[0] : unique;
};

const parseManualResponsibilityInput = (input: string): PersonSelection[] => {
  const tokens = extractStringTokens(input);
  const seen = new Set<string>();
  const selections: PersonSelection[] = [];
  tokens.forEach(token => {
    const key = token.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    const containsAt = token.indexOf('@') !== -1;
    selections.push({
      displayName: token,
      email: containsAt ? token : undefined,
      loginName: containsAt ? token : undefined
    });
  });
  return selections;
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
    const idCandidates = extractNumberTokens(initialValues?.RcResponsibilityId);
    const emailCandidates = extractStringTokens(initialValues?.RcResponsibilityEmail);
    const loginCandidates = extractStringTokens(initialValues?.RcResponsibilityLoginName);
    const nameCandidates = extractStringTokens(initialValues?.RcResponsibility);

    const maxLen = Math.max(
      idCandidates.length,
      emailCandidates.length,
      loginCandidates.length,
      nameCandidates.length
    );

    if (maxLen === 0) {
      if (typeof initialValues?.RcResponsibility === 'string' && initialValues.RcResponsibility.trim().length) {
        return [{ displayName: initialValues.RcResponsibility.trim() }];
      }
      return [];
    }

    const selections: PersonSelection[] = [];
    for (let index = 0; index < maxLen; index += 1) {
      const numericId = idCandidates[index];
      const email = emailCandidates[index];
      const loginName = loginCandidates[index] ?? emailCandidates[index];
      const displayName = nameCandidates[index] ?? email ?? loginName;

      if (
        (numericId !== undefined && !isNaN(numericId)) ||
        (email && email.length > 0) ||
        (loginName && loginName.length > 0) ||
        (displayName && displayName.length > 0)
      ) {
        selections.push({
          id: typeof numericId === 'number' && !isNaN(numericId) ? numericId : undefined,
          email,
          loginName,
          displayName
        });
      }
    }

    if (selections.length > 0) {
      return selections;
    }

    if (typeof initialValues?.RcResponsibility === 'string' && initialValues.RcResponsibility.trim().length) {
      return [{ displayName: initialValues.RcResponsibility.trim() }];
    }
    return [];
  }, [initialValues?.RcResponsibility, initialValues?.RcResponsibilityEmail, initialValues?.RcResponsibilityId, initialValues?.RcResponsibilityLoginName]);

  const createInitialState = useCallback((): ReusableComponentsFormState => {
    const responsibility = mapInitialResponsibility();
    const responsibilityDisplay = displayFromSelection(responsibility);
    const collapsedIds = collapseNumberValues(responsibility.map(person => person.id));
    const collapsedEmails = collapseStringValues(responsibility.map(person => person.email));
    const collapsedLogins = collapseStringValues(
      responsibility.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    const fallbackIds = collapsedIds ?? collapseNumberValues(extractNumberTokens(initialValues?.RcResponsibilityId));
    const fallbackEmails = collapsedEmails ?? collapseStringValues(extractStringTokens(initialValues?.RcResponsibilityEmail));
    const fallbackLogins = collapsedLogins ?? collapseStringValues(
      extractStringTokens(
        initialValues?.RcResponsibilityLoginName ??
        initialValues?.RcResponsibilityEmail ??
        initialValues?.RcResponsibility
      )
    );

    return {
      RcComponentName: initialValues?.RcComponentName ?? '',
      RcLocation: initialValues?.RcLocation ?? '',
      RcPurposeMainFunctionality: initialValues?.RcPurposeMainFunctionality ?? '',
      RcResponsibility: responsibilityDisplay,
      RcResponsibilityId: fallbackIds,
      RcResponsibilityEmail: fallbackEmails,
      RcResponsibilityLoginName: fallbackLogins ?? fallbackEmails,
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

    const display = displayFromSelection(mapped);
    const responsibilityIds = collapseNumberValues(mapped.map(person => person.id));
    const responsibilityEmails = collapseStringValues(mapped.map(person => person.email));
    const responsibilityLogins = collapseStringValues(
      mapped.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    setFormState(prev => ({
      ...prev,
      RcResponsibility: display,
      RcResponsibilityId: responsibilityIds,
      RcResponsibilityEmail: responsibilityEmails,
      RcResponsibilityLoginName: responsibilityLogins ?? responsibilityEmails
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
    if (!trimmed) {
      setFormState(prev => ({
        ...prev,
        RcResponsibility: '',
        RcResponsibilityId: undefined,
        RcResponsibilityEmail: undefined,
        RcResponsibilityLoginName: undefined
      }));
      setSelectedResponsibility([]);
      return;
    }

    const parsedSelections = parseManualResponsibilityInput(trimmed);
    const hasParsed = parsedSelections.length > 0;
    const effectiveSelections = hasParsed
      ? parsedSelections
      : [{
        displayName: trimmed,
        email: trimmed.indexOf('@') !== -1 ? trimmed : undefined,
        loginName: trimmed.indexOf('@') !== -1 ? trimmed : undefined
      }];
    const display = displayFromSelection(effectiveSelections);
    const emailValues = collapseStringValues(effectiveSelections.map(person => person.email));
    const loginValues = collapseStringValues(
      effectiveSelections.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    setFormState(prev => ({
      ...prev,
      RcResponsibility: display,
      RcResponsibilityId: undefined,
      RcResponsibilityEmail: emailValues,
      RcResponsibilityLoginName: loginValues ?? emailValues ?? display
    }));

    setSelectedResponsibility(effectiveSelections);
    if (errors.RcResponsibility) {
      setErrors(prev => ({ ...prev, RcResponsibility: '' }));
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

    const responsibilityIds = collapseNumberValues(responsibilitySnapshot.map(person => person.id));
    const responsibilityEmails = collapseStringValues(responsibilitySnapshot.map(person => person.email));
    const responsibilityLogins = collapseStringValues(
      responsibilitySnapshot.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    onSubmit?.({
      RcComponentName: nextState.RcComponentName.trim(),
      RcLocation: nextState.RcLocation.trim(),
      RcPurposeMainFunctionality: nextState.RcPurposeMainFunctionality.trim(),
      RcResponsibility: nextState.RcResponsibility.trim(),
      RcResponsibilityId: responsibilityIds,
      RcResponsibilityEmail: responsibilityEmails,
      RcResponsibilityLoginName: responsibilityLogins ?? responsibilityEmails,
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

  const defaultSelectedUsers = uniqueStrings(
    selectedResponsibility.map(person => person.email || person.loginName || person.displayName)
  );
  const personSelectionLimit = Math.max(
    Math.max(defaultSelectedUsers.length, selectedResponsibility.length) + 5,
    5
  );

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
                  personSelectionLimit={personSelectionLimit}
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