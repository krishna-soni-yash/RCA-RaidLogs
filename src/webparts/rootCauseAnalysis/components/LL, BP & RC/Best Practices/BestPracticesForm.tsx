/*eslint-disable*/
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
  BpResponsibilityId?: number | number[];
  BpResponsibilityEmail?: string | string[];
  BpResponsibilityLoginName?: string | string[];
  BpRemarks: string;
  DataType: string;
};

type BestPracticesFormErrors = {
  BpBestPracticesDescription: string;
  BpReferences: string;
  BpResponsibility: string;
  BpRemarks: string;
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
  const [peoplePickerFailed, setPeoplePickerFailed] = useState<boolean>(false);
  const isReadOnly = mode === 'view';
  const peoplePickerWebUrl = useMemo(() => {
    return (context as any)?.pageContext?.web?.absoluteUrl ?? window.location.origin;
  }, [context]);

  const mapInitialResponsibility = useCallback((): PersonSelection[] => {
    const idCandidates = extractNumberTokens(initialValues?.BpResponsibilityId);
    const emailCandidates = extractStringTokens(initialValues?.BpResponsibilityEmail);
    const loginCandidates = extractStringTokens(initialValues?.BpResponsibilityLoginName);
    const nameCandidates = extractStringTokens(initialValues?.BpResponsibility);

    const maxLen = Math.max(
      idCandidates.length,
      emailCandidates.length,
      loginCandidates.length,
      nameCandidates.length
    );

    if (maxLen === 0) {
      if (typeof initialValues?.BpResponsibility === 'string' && initialValues.BpResponsibility.trim().length) {
        return [{ displayName: initialValues.BpResponsibility.trim() }];
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

    if (typeof initialValues?.BpResponsibility === 'string' && initialValues.BpResponsibility.trim().length) {
      return [{ displayName: initialValues.BpResponsibility.trim() }];
    }
    return [];
  }, [initialValues?.BpResponsibility, initialValues?.BpResponsibilityEmail, initialValues?.BpResponsibilityId, initialValues?.BpResponsibilityLoginName]);

  const createInitialState = useCallback((): BestPracticesFormState => {
    const responsibility = mapInitialResponsibility();
    const responsibilityDisplay = displayFromSelection(responsibility);
    const collapsedIds = collapseNumberValues(responsibility.map(person => person.id));
    const collapsedEmails = collapseStringValues(responsibility.map(person => person.email));
    const collapsedLogins = collapseStringValues(
      responsibility.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    const fallbackIds = collapsedIds ?? collapseNumberValues(extractNumberTokens(initialValues?.BpResponsibilityId));
    const fallbackEmails = collapsedEmails ?? collapseStringValues(extractStringTokens(initialValues?.BpResponsibilityEmail));
    const fallbackLogins = collapsedLogins ?? collapseStringValues(
      extractStringTokens(
        initialValues?.BpResponsibilityLoginName ??
        initialValues?.BpResponsibilityEmail ??
        initialValues?.BpResponsibility
      )
    );

    return {
      BpBestPracticesDescription: initialValues?.BpBestPracticesDescription ?? '',
      BpReferences: initialValues?.BpReferences ?? '',
      BpResponsibility: responsibilityDisplay,
      BpResponsibilityId: fallbackIds,
      BpResponsibilityEmail: fallbackEmails,
      BpResponsibilityLoginName: fallbackLogins ?? fallbackEmails,
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
    setPeoplePickerFailed(false);

    const display = displayFromSelection(mapped);
    const responsibilityIds = collapseNumberValues(mapped.map(person => person.id));
    const responsibilityEmails = collapseStringValues(mapped.map(person => person.email));
    const responsibilityLogins = collapseStringValues(
      mapped.map(person => person.loginName ?? person.email ?? person.displayName)
    );

    setFormState(prev => ({
      ...prev,
      BpResponsibility: display,
      BpResponsibilityId: responsibilityIds,
      BpResponsibilityEmail: responsibilityEmails,
      BpResponsibilityLoginName: responsibilityLogins ?? responsibilityEmails
    }));

    if (mapped.length > 0 && errors.BpResponsibility) {
      setErrors(prev => ({ ...prev, BpResponsibility: '' }));
    }
  }, [errors.BpResponsibility, isReadOnly]);

  const handleManualResponsibilityInput = useCallback((_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
    if (isReadOnly) {
      return;
    }

    const trimmed = (value ?? '').trim();
    if (!trimmed) {
      setFormState(prev => ({
        ...prev,
        BpResponsibility: '',
        BpResponsibilityId: undefined,
        BpResponsibilityEmail: undefined,
        BpResponsibilityLoginName: undefined
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
      BpResponsibility: display,
      BpResponsibilityId: undefined,
      BpResponsibilityEmail: emailValues,
      BpResponsibilityLoginName: loginValues ?? emailValues ?? display
    }));

    setSelectedResponsibility(effectiveSelections);
    if (errors.BpResponsibility) {
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

      const responsibilityIds = collapseNumberValues(responsibilitySnapshot.map(person => person.id));
      const responsibilityEmails = collapseStringValues(responsibilitySnapshot.map(person => person.email));
      const responsibilityLogins = collapseStringValues(
        responsibilitySnapshot.map(person => person.loginName ?? person.email ?? person.displayName)
      );

      onSubmit?.({
        BpBestPracticesDescription: nextState.BpBestPracticesDescription.trim(),
        BpReferences: nextState.BpReferences.trim(),
        BpResponsibility: nextState.BpResponsibility.trim(),
        BpResponsibilityId: responsibilityIds,
        BpResponsibilityEmail: responsibilityEmails,
        BpResponsibilityLoginName: responsibilityLogins ?? responsibilityEmails,
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
            {!peoplePickerFailed && context ? (
              <PeoplePickerErrorBoundary onError={() => setPeoplePickerFailed(true)}>
                <PeoplePicker
                  key={`bp-responsibility-${peoplePickerKey}`}
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
                  value={formState.BpResponsibility}
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