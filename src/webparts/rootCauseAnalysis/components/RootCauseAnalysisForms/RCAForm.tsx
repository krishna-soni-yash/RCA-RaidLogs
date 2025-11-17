import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  DefaultButton,
  PrimaryButton,
  Pivot,
  PivotItem,
  NormalPeoplePicker,
  IPersonaProps,
  IBasePickerSuggestionsProps
} from '@fluentui/react';
import { saveRCAItem, updateRCAItem } from '../../../../repositories/repositoriesInterface/RCARepository';
import { WebPartContext } from '@microsoft/sp-webpart-base';
interface RCAFormProps {
  onSubmit?: (data: any) => void;
  initialData?: any;
  context?: WebPartContext;
}

export default function RCAForm({ onSubmit, initialData,context }: RCAFormProps) {
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

  // --- people picker helpers ---
  const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested people',
    mostRecentlyUsedHeaderText: 'Recent',
    noResultsFoundText: 'No results found',
    showRemoveButtons: false
  };

  // simple resolver - allow freeform entry; set secondaryText when input looks like an email
  const onResolveSuggestions = (filter: string, _selected?: IPersonaProps[] | undefined): IPersonaProps[] => {
    if (!filter || filter.trim().length === 0) return [];
    const trimmed = filter.trim();
    const isEmail = /\S+@\S+\.\S+/.test(trimmed);
    return [{
      key: trimmed,
      primaryText: trimmed,
      text: trimmed,
      secondaryText: isEmail ? trimmed : undefined
    }];
  };

  // store responsibility as array of strings (emails or names) so repository can accept string[] or string
  const handlePeoplePickerChange = (actionKey: string) => (items?: IPersonaProps[] | undefined) => {
    const values = (items || []).map(p => (p.secondaryText || p.text || p.primaryText)).filter(Boolean);
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

      <Dropdown
        label="Type of Action"
        options={actionTypeOptions}
        multiSelect
        // selectedKeys expects string[] for multi-select
        selectedKeys={form.actionType && form.actionType.length ? form.actionType : undefined}
        onChange={(_, o) => {
          const key = o?.key as string;
          const current: string[] = form.actionType || [];
          const next = current.indexOf(key) !== -1 ? current.filter(k => k !== key) : [...current, key];
          update('actionType', next);
        }}
        disabled={actionTypeDisabled}
      />

      {/* show selected action types */}
      <div style={{ fontSize: 13, color: '#605e5c' }}>
        Selected: {(form.actionType && form.actionType.length) ? (form.actionType as string[]).join(', ') : 'None'}
      </div>

      {/* Tabs for each selected action type; fallback single fields when none selected */}
      {form.actionType && form.actionType.length > 0 ? (
        <Pivot aria-label="Action type tabs">
          {(form.actionType as string[]).map((act) => (
            <PivotItem headerText={act} key={act}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 8 }}>
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
                  <NormalPeoplePicker
                    onResolveSuggestions={onResolveSuggestions}
                    onChange={handlePeoplePickerChange(act)}
                    pickerSuggestionsProps={suggestionProps}
                    selectedItems={
                      // support both stored array or semicolon-separated string from older saves
                      (() => {
                        const raw = actionDetails[act]?.responsibility;
                        const values: string[] = Array.isArray(raw)
                          ? raw
                          : (typeof raw === 'string' && raw.length ? raw.split(/; ?/).map((s: string) => s.trim()).filter(Boolean) : []);
                        return values.map((s: string) => {
                          const isEmail = /\S+@\S+\.\S+/.test(s);
                          return { key: s, primaryText: s, text: s, secondaryText: isEmail ? s : undefined };
                        });
                      })()
                    }
                    resolveDelay={300}
                    inputProps={{ 'aria-label': 'Responsibility' }}
                  />
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