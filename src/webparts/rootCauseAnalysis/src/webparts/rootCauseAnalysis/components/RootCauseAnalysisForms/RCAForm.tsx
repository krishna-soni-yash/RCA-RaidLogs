import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  TextField,
  Text,
  Dropdown,
  IDropdownOption,
  DatePicker,
  DefaultButton,
  PrimaryButton,
  Pivot,
  PivotItem,
  Checkbox // added
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { NormalPeoplePicker, IPersonaProps } from '@fluentui/react';
import { saveRCAItem, updateRCAItem } from '../../../../../../../repositories/RCARepository';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from '../../../../../components/RootCauseAnalysis.module.scss';
interface RCAFormProps {
  onSubmit?: (data: any) => void;
  initialData?: any;
  context?: WebPartContext;
}

// small ErrorBoundary to catch PeoplePicker runtime errors (e.g. PeopleSearchService failures)
class PeoplePickerErrorBoundary extends React.Component<{ onError?: () => void }, { hasError: boolean }> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() {
    return { hasError: true };
  }
  componentDidCatch(error: any, info: any) {
    console.error('PeoplePicker render error caught by boundary', error, info);
    if (this.props.onError) this.props.onError();
  }
  render() {
    if (this.state.hasError) return null; // parent component will show fallback
    return this.props.children;
  }
}

export default function RCAForm({ onSubmit, initialData, context }: RCAFormProps) {
  const [form, setForm] = useState<any>({
    problemStatement: initialData?.problemStatement || '',
    causeCategory: initialData?.causeCategory || '',
    source: initialData?.source || '',
    priority: initialData?.priority || '',
    relatedMetric: initialData?.relatedMetric || '',
    causes: initialData?.causes || '',
    rootCauses: initialData?.rootCauses || '',
    analysisTechnique: initialData?.analysisTechnique || '',
    // store actionType as an array of selected keys
    actionType: initialData?.actionType
      ? (Array.isArray(initialData.actionType) ? initialData.actionType : [initialData.actionType])
      : [],
    actionPlan: initialData?.actionPlan || '',
    responsibility: initialData?.responsibility || '',
    plannedClosureDate: initialData?.plannedClosureDate ? new Date(initialData.plannedClosureDate) : undefined,
    actualClosureDate: initialData?.actualClosureDate ? new Date(initialData.actualClosureDate) : undefined,
    performanceBefore: initialData?.performanceBefore || '',
    performanceAfter: initialData?.performanceAfter || '',
    quantitativeEffectiveness: initialData?.quantitativeEffectiveness || '',
    remarks: initialData?.remarks || '',
    attachments: initialData?.attachments || []
  });

  // disabled state for the Type of Action dropdown
  const [actionTypeDisabled, setActionTypeDisabled] = useState<boolean>(false);

  // per-action-type details (actionPlan/responsibility/dates)
  const [actionDetails, setActionDetails] = useState<Record<string, any>>(initialData?.actionDetails || {});

  const causeCategoryOptions: IDropdownOption[] = [
    { key: 'Special', text: 'Special' },
    { key: 'Common', text: 'Common' }
  ];

  const sourceOptions: IDropdownOption[] = [
    { key: 'Audit Findings', text: 'Audit Findings' },
    { key: 'Metrics', text: 'Metrics' },
    { key: 'Review Findings', text: 'Review Findings' },
    { key: 'Testing Findings', text: 'Testing Findings' },
    { key: 'Customer Feedback', text: 'Customer Feedback' }
  ];

  const priorityOptions: IDropdownOption[] = [
    { key: 'High', text: 'High' },
    { key: 'Medium', text: 'Medium' },
    { key: 'Low', text: 'Low' }
  ];

  const actionTypeOptions: IDropdownOption[] = [
    { key: 'Correction', text: 'Correction' },
    { key: 'Corrective Action', text: 'Corrective Action' },
    { key: 'Preventive Action', text: 'Preventive Action' }
  ];

  const update = (key: string, value: any) => setForm((s: any) => ({ ...s, [key]: value }));

  // new: control whether the attachments panel is expanded
  const [attachmentsOpen, setAttachmentsOpen] = useState<boolean>(false);

  const onFilesAdded = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.prototype.slice.call(e.target.files) : [];
    if (files.length === 0) return;
    setForm((s: any) => ({ ...s, attachments: [...(s.attachments || []), ...files] }));
    e.currentTarget.value = '';
  };

  const removeAttachment = (index: number) => {
    setForm((s: any) => ({ ...s, attachments: (s.attachments || []).filter((_: any, i: number) => i !== index) }));
  };

  // ensure actionDetails has entries for selected action types
  useEffect(() => {
    const selected: string[] = form.actionType || [];
    if (!selected || selected.length === 0) return;
    setActionDetails((prev) => {
      const next = { ...prev };
      selected.forEach((k) => {
        if (!next[k]) {
          next[k] = {
            actionPlan: '',
            responsibility: '',
            plannedClosureDate: undefined,
            actualClosureDate: undefined
          };
        }
      });
      return next;
    });
  }, [form.actionType]);

  const updateActionDetail = (actionKey: string, field: string, value: any) => {
    setActionDetails((prev) => ({
      ...prev,
      [actionKey]: {
        ...(prev[actionKey] || {}),
        [field]: value
      }
    }));
  };





  // store responsibility as array of strings (email/login/name) so repository can accept string[] or string
  const handlePeoplePickerChange = (actionKey: string) => (items: any[]) => {
    const list = Array.isArray(items) ? items : [];
    const values = list
      .map((p: any) => p?.secondaryText || p?.loginName || p?.text || p?.id)
      .filter((v: any): v is string => typeof v === 'string' && v.length > 0);
    updateActionDetail(actionKey, 'responsibility', values);
  };
  // --- end people picker helpers ---

  const onSave = async () => {
    // keep top-level fields in sync with first selected action type for compatibility
    const firstKey = (form.actionType && form.actionType.length) ? form.actionType[0] : undefined;
    const payload = {
      ...form,
      actionDetails
    };
    if (firstKey && actionDetails[firstKey]) {
      payload.actionPlan = actionDetails[firstKey].actionPlan;
      payload.responsibility = actionDetails[firstKey].responsibility;
      payload.plannedClosureDate = actionDetails[firstKey].plannedClosureDate;
      payload.actualClosureDate = actionDetails[firstKey].actualClosureDate;
    }
    console.log('Prepared payload', payload);
    //setPayloadState(payload);

    // build item matching IRCAList fields expected by saveRCAItem
    const formatDate = (d: any) => {
      if (!d) return undefined;
      if (d instanceof Date && !isNaN(d.getTime())) return d.toISOString();
      return d;
    };

    const item: any = {};
    // title / problem statement
    item.LinkTitle = form.problemStatement || '';
    //item.ProblemStatement = form.problemStatement || '';

    // top-level mappings
    item.CauseCategory = form.causeCategory || '';
    item.RCASource = form.source || '';
    item.RCAPriority = form.priority || '';
    item.RelatedMetric = form.relatedMetric || '';
    item.Cause = form.causes || '';
    item.RootCause = form.rootCauses || '';
    item.RCATechniqueUsedAndReference = form.analysisTechnique || '';
    // join action types into a single string similar to existing list storage
    item.RCATypeOfAction = (form.actionType && form.actionType.length) ? (form.actionType as string[]).join(', ') : '';

    // map per-action-type details to repository fields using suffix mapping
    Object.keys(actionDetails || {}).forEach((actKey) => {
      const details = actionDetails[actKey] || {};
      // determine suffix used in repository field names
      let suffix = '';
      if (actKey.toLowerCase().indexOf('correction') !== -1) suffix = 'Correction';
      else if (actKey.toLowerCase().indexOf('corrective') !== -1) suffix = 'Corrective';
      else if (actKey.toLowerCase().indexOf('preventive') !== -1) suffix = 'Preventive';
      else {
        // fallback: sanitize actKey to use as suffix (remove spaces)
        suffix = actKey.replace(/\s+/g, '');
      }

      if (details.actionPlan !== undefined) item[`ActionPlan${suffix}`] = details.actionPlan;
      if (details.responsibility !== undefined) item[`Responsibility${suffix}`] = details.responsibility;
      if (details.plannedClosureDate !== undefined) item[`PlannedClosureDate${suffix}`] = formatDate(details.plannedClosureDate);
      if (details.actualClosureDate !== undefined) item[`ActualClosureDate${suffix}`] = formatDate(details.actualClosureDate);
    });

    if (form.performanceBefore !== undefined) item.PerformanceBeforeActionPlan = form.performanceBefore;
    if (form.performanceAfter !== undefined) item.PerformanceAfterActionPlan = form.performanceAfter;
    if (form.quantitativeEffectiveness !== undefined) item.QuantitativeOrStatisticalEffecti = form.quantitativeEffectiveness;
    if (form.remarks !== undefined) item.Remarks = form.remarks;

    // determine repository id from initialData when editing
    const repoId = initialData ? (initialData.__repoId ?? initialData.__id ?? initialData.id ?? initialData.ID) : undefined;
    const numericRepoId = repoId ? Number(repoId) : undefined;

    // call repository save/update
    try {
      if (context) {
        if (numericRepoId && numericRepoId > 0) {
          await updateRCAItem(numericRepoId, item, context);
          console.log('RCA updated', numericRepoId);
        } else {
          const result = await saveRCAItem(item, context);
          console.log('RCA saved', result);
        }
      } else {
        console.warn('No WebPart context provided to save/update RCA item - skipping backend call.');
      }

      if (onSubmit) onSubmit(payload);
      else console.log('RCA Form submit', payload);
    } catch (err) {
      console.error('Failed to save/update RCA item', err);
    }
  };

  const onReset = () => {
    setForm({
      problemStatement: '',
      causeCategory: '',
      source: '',
      priority: '',
      relatedMetric: '',
      causes: '',
      rootCauses: '',
      analysisTechnique: '',
      // reset actionType to empty array
      actionType: [],
      actionPlan: '',
      responsibility: '',
      plannedClosureDate: undefined,
      actualClosureDate: undefined,
      performanceBefore: '',
      performanceAfter: '',
      quantitativeEffectiveness: '',
      remarks: '',
      // ensure attachments cleared on reset
      attachments: []
    });
    setActionDetails({});
    // ensure dropdown is enabled after reset
    setActionTypeDisabled(false);
  };

  // when Cause Category is "Special", pre-populate and disable Type of Action
  useEffect(() => {
    if (form.causeCategory === 'Special') {
      update('actionType', ['Correction']);
      setActionTypeDisabled(true);
    } else {
      setActionTypeDisabled(false);
    }
  }, [form.causeCategory]);

  // track if PeoplePicker failed so we can fall back
  const [peoplePickerFailed, setPeoplePickerFailed] = useState<boolean>(false);

  // debug helper - log context & environment
  useEffect(() => {
    if (context) {
      try {
        // small sanity logs to help debugging PeopleSearchService failures
        // check if running in workbench (suggests suggestions won't work)
        const isWorkbench = (context as any)?.pageContext === undefined || window.location.hostname.indexOf('localhost') !== -1;
        console.info('RCAForm PeoplePicker context:', {
          webAbsoluteUrl: (context as any)?.pageContext?.web?.absoluteUrl,
          isWorkbench
        });
      } catch (e) {
        console.warn('Unable to inspect context for PeoplePicker debug info', e);
      }
    } else {
      console.info('RCAForm: no WebPart context passed; PeoplePicker will be unavailable.');
    }
  }, [context]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <TextField
        label="Problem statement (Causal Analysis Trigger)"
        value={form.problemStatement}
        onChange={(_, v) => update('problemStatement', v)}
      />

      <div style={{ display: 'flex', gap: 12 }}>
        <div style={{ flex: 1 }}>
          <Dropdown
            label="Cause Category"
            options={causeCategoryOptions}
            selectedKey={form.causeCategory || undefined}
            onChange={(_, o) => update('causeCategory', o?.key)}
          />
        </div>

        <div style={{ flex: 1 }}>
          <Dropdown
            label="Source"
            options={sourceOptions}
            selectedKey={form.source || undefined}
            onChange={(_, o) => update('source', o?.key)}
          />
        </div>

        <div style={{ flex: 1 }}>
          <Dropdown
            label="Priority"
            options={priorityOptions}
            selectedKey={form.priority || undefined}
            onChange={(_, o) => update('priority', o?.key)}
          />
        </div>
      </div>

      {/* Related Metric (attachments link removed — attachments moved to bottom) */}
      <div>
        <TextField
          label="Related Metric (if any)"
          value={form.relatedMetric}
          onChange={(_, v) => update('relatedMetric', v)}
        />
      </div>

      <TextField
        label="Cause(s)"
        value={form.causes}
        onChange={(_, v) => update('causes', v)}
      />

      <TextField
        label="Root Cause(s)"
        value={form.rootCauses}
        onChange={(_, v) => update('rootCauses', v)}
      />

      <TextField
        label="Root Cause Analysis Technique Used and Reference (if any)"
        value={form.analysisTechnique}
        onChange={(_, v) => update('analysisTechnique', v)}
        multiline
        rows={3}
      />

      {/* Type of Action — replaced multiSelect Dropdown with checkboxes */}
      <div>
        <label style={{ display: 'block', marginBottom: 6, fontSize: 12, color: '#605e5c' }}>Type of Action</label>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 8 }}>
          {actionTypeOptions.map(opt => {
            const key = opt.key as string;
            const checked = Array.isArray(form.actionType) && form.actionType.indexOf(key) !== -1;
            return (
              <Checkbox
                key={key}
                label={opt.text}
                checked={checked}
                onChange={(_, isChecked) => {
                  const current: string[] = form.actionType || [];
                  const next = isChecked ? [...current.filter(k => k !== key), key] : current.filter(k => k !== key);
                  update('actionType', next);
                }}
                disabled={actionTypeDisabled}
              />
            );
          })}
        </div>
      </div>

      {/* show selected action types */}
      {/* <div style={{ fontSize: 13, color: '#605e5c' }}>
        Selected: {(form.actionType && form.actionType.length) ? (form.actionType as string[]).join(', ') : 'None'}
      </div> */}

      {/* Tabs for each selected action type; fallback single fields when none selected */}
      {form.actionType && form.actionType.length > 0 ? (
        <Pivot aria-label="Action type tabs" style={{ marginTop: 16 }}>
          {(form.actionType as string[]).map((act) => (
            <PivotItem headerText={act} key={act} >
              <div style={{ padding: '16px', background: '#fefefe' }}>
                <div className={styles.actionRow}>
                  <div className={styles.actionHeader}>
                    <Text variant="mediumPlus">{act}</Text>
                  </div>
                  {/* <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 8 }}> */}
                    <TextField
                      label="Action Plan"
                      value={actionDetails[act]?.actionPlan || ''}
                      onChange={(_, v) => updateActionDetail(act, 'actionPlan', v)}
                      multiline
                      rows={4}
                    />

                    {/* Responsibility converted to People Picker */}
                    <div>
                      <label style={{ display: 'block', marginBottom: 6, color: '#605e5c', fontSize: 12 }}>Responsibility</label>

                      {/* Wrap PeoplePicker in an ErrorBoundary; on error set peoplePickerFailed */}
                      {!peoplePickerFailed && context ? (
                        <PeoplePickerErrorBoundary onError={() => setPeoplePickerFailed(true)}>
                          <PeoplePicker
                            context={
                              {
                                ...(context as any),
                                absoluteUrl: (context as any)?.pageContext?.web?.absoluteUrl ?? window.location.origin
                              } as any
                            }
                            titleText=""
                            personSelectionLimit={5}
                            showtooltip
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={300}
                            ensureUser={true}
                            placeholder="Type a name or email..."
                            defaultSelectedUsers={
                              (() => {
                                const raw = actionDetails[act]?.responsibility;
                                if (!raw) return [];
                                if (Array.isArray(raw)) {
                                  return raw
                                    .map((r: any) =>
                                      typeof r === 'string'
                                        ? r
                                        : r.secondaryText || r.loginName || r.text || ''
                                    )
                                    .filter(Boolean);
                                }
                                if (typeof raw === 'string' && raw.length) {
                                  return raw
                                    .split(/; ?/)
                                    .map((s: string) => s.trim())
                                    .filter(Boolean);
                                }
                                return [];
                              })()
                            }
                            onChange={handlePeoplePickerChange(act)}
                          />
                        </PeoplePickerErrorBoundary>
                      ) : (
                        // fallback: freeform NormalPeoplePicker so user can still enter assignees
                        <>
                          <div style={{ marginBottom: 6, color: '#a19f9d', fontSize: 12 }}>
                            {peoplePickerFailed ? 'People search failed — use manual entry.' : 'PeoplePicker unavailable.'}
                          </div>
                          <NormalPeoplePicker
                            onResolveSuggestions={(filterText: string, _selected?: IPersonaProps[]) => {
                              if (!filterText || filterText.trim().length === 0) return [];
                              const t = filterText.trim();
                              const isEmail = /\S+@\S+\.\S+/.test(t);
                              return [{
                                key: t,
                                primaryText: t,
                                text: t,
                                secondaryText: isEmail ? t : undefined
                              }];
                            }}
                            onChange={(items?: IPersonaProps[] | undefined) => {
                              const values = (items || []).map(p => (p.secondaryText || p.text || p.primaryText)).filter(Boolean);
                              updateActionDetail(act, 'responsibility', values);
                            }}
                            selectedItems={
                              (() => {
                                const raw = actionDetails[act]?.responsibility;
                                const values: IPersonaProps[] = [];
                                if (!raw) return values;
                                const arr = Array.isArray(raw) ? raw : (typeof raw === 'string' ? raw.split(/; ?/).map((s: string) => s.trim()).filter(Boolean) : []);
                                return arr.map((s: string) => {
                                  const isEmail = /\S+@\S+\.\S+/.test(s);
                                  return { key: s, primaryText: s, text: s, secondaryText: isEmail ? s : undefined } as IPersonaProps;
                                });
                              })()
                            }
                            resolveDelay={300}
                            inputProps={{ 'aria-label': 'Responsibility' }}
                          />
                        </>
                      )}
                    </div>

                    <div style={{ display: 'flex', gap: 12 }}>
                      <div style={{ flex: 1 }}>
                        <DatePicker
                          label="Planned Closure Date"
                          isMonthPickerVisible={false}
                          value={actionDetails[act]?.plannedClosureDate}
                          onSelectDate={(d) => updateActionDetail(act, 'plannedClosureDate', d)}
                        />
                      </div>
                      <div style={{ flex: 1 }}>
                        <DatePicker
                          label="Actual Closure Date"
                          isMonthPickerVisible={false}
                          value={actionDetails[act]?.actualClosureDate}
                          onSelectDate={(d) => updateActionDetail(act, 'actualClosureDate', d)}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              {/* </div> */}

            </PivotItem>
          ))}
        </Pivot>
      ) : (
        // fallback single set of fields (keeps previous behavior)
        <></>
      )}

      <TextField
        label="Performance before action plan"
        value={form.performanceBefore}
        onChange={(_, v) => update('performanceBefore', v)}
        multiline
        rows={3}
      />

      <TextField
        label="Performance after action plan"
        value={form.performanceAfter}
        onChange={(_, v) => update('performanceAfter', v)}
        multiline
        rows={3}
      />

      <TextField
        label="Quantitative / Statistical effectiveness"
        value={form.quantitativeEffectiveness}
        onChange={(_, v) => update('quantitativeEffectiveness', v)}
        multiline
        rows={3}
      />

      <TextField
        label="Remarks"
        value={form.remarks}
        onChange={(_, v) => update('remarks', v)}
        multiline
        rows={2}
      />

      {/* Attachments area moved to the bottom of the form */}
      <div style={{ marginTop: 6 }}>
        <a
          role="button"
          aria-expanded={attachmentsOpen}
          onClick={() => setAttachmentsOpen((s) => !s)}
          style={{
            fontSize: 13,
            cursor: 'pointer',
            color: '#605e5c',
            textDecoration: 'underline',
            display: 'inline-block',
            marginBottom: 8
          }}
        >
          Attachments{form.attachments && form.attachments.length ? ` (${form.attachments.length})` : ''}
        </a>

        {attachmentsOpen && (
          <div style={{ border: '1px solid #e1dfdd', padding: 8, borderRadius: 4 }}>
            <input
              type="file"
              multiple
              onChange={onFilesAdded}
              style={{ marginBottom: 8 }}
              aria-label="Add attachments"
            />
            {(form.attachments && form.attachments.length > 0) && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {form.attachments.map((f: File, idx: number) => (
                  <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 420 }}>{f.name}</span>
                    <DefaultButton text="Remove" onClick={() => removeAttachment(idx)} />
                  </div>
                ))}
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 8 }}>
              <a onClick={() => setAttachmentsOpen(false)} style={{ cursor: 'pointer', textDecoration: 'underline', color: '#605e5c' }}>Close</a>
            </div>
          </div>
        )}
      </div>

      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
        <DefaultButton text="Reset" onClick={onReset} />
        <PrimaryButton text="Save" onClick={onSave} />
      </div>
    </div>
  );
}