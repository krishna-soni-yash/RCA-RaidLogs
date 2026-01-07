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
  Checkbox // added
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

import { saveRCAItem, updateRCAItem, deleteRCAAttachment, uploadRCAAttachment } from '../../../../repositories/RCARepository';
import { MessageModal, MessageType } from '../ModalPopups'; // adjust path if your MessageModal lives elsewhere

import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from '../../components/RootCauseAnalysis.module.scss';
import { getMetricsFromProjectMetrics, getSubMetricsFromProjectMetrics } from '../../../../repositories/MetricsRepository';
import { MetricsRepository } from '../../../../repositories/MetricsRepository';
import IProjectMetricsRepository from '../../../../repositories/repositoriesInterface/IProjectMetricsRepository';
import { GenericService } from '../../../../services/GenericServices';
import IGenericService from '../../../../services/IGenericServices';
//import { SPHttpClient } from '@microsoft/sp-http'; // <-- added
//import { SubSiteListNames } from '../../../../common/Constants';

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
  const [MetricsData, setMetricsData] = React.useState<Array<{ key: string; text: string }>>([]);
  const [SubMetricsData, setSubMetricsData] = React.useState<Array<{ key: string; text: string }>>([]);
  // modal state for save/update success message
  const [showMessageModal, setShowMessageModal] = React.useState<boolean>(false);
  const [messageText, setMessageText] = React.useState<string>('');
  const [messageType, setMessageType] = React.useState<MessageType>('info');
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

  //  const relatedMetricOptions: IDropdownOption[] = [
  //   { key: 'High', text: 'High' },
  //   { key: 'Medium', text: 'Medium' },
  //   { key: 'Low', text: 'Low' }
  // ];
  const actionTypeOptions: IDropdownOption[] = [
    { key: 'Correction', text: 'Correction' },
    { key: 'Corrective Action', text: 'Corrective Action' },
    { key: 'Preventive Action', text: 'Preventive Action' }
  ];

  // add validation error state
  const [errors, setErrors] = useState<Record<string, string>>({});
  const showMessage = (message: string, type: MessageType): void => {
    setMessageText(message);
    setMessageType(type);
    setShowMessageModal(true);
  };

  const handleDismissMessage = (): void => {
    setShowMessageModal(false);
  };

  // const handleValidationError = (message: string): void => {
  //   showMessage(message, 'warning');
  // };
  // update helper - clear field error when value becomes non-empty
  const update = (key: string, value: any) => {
    setForm((s: any) => ({ ...s, [key]: value }));
    setErrors(prev => {
      const next = { ...prev };
      if (value !== null && value !== undefined && value !== '') delete next[key];
      return next;
    });
  };

  // new: control whether the attachments panel is expanded
  const [attachmentsOpen, setAttachmentsOpen] = useState<boolean>(false);

  const onFilesAdded = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.prototype.slice.call(e.target.files) : [];
    if (files.length === 0) return;
    setForm((s: any) => ({ ...s, attachments: [...(s.attachments || []), ...files] }));
    e.currentTarget.value = '';
  };

  const removeAttachment = (index: number) => {
    setForm((s: any) => {
      const current = s.attachments || [];
      const item = current[index];
      if (item instanceof File) {
        return { ...s, attachments: current.filter((_: any, i: number) => i !== index) };
      }
      const fileName = (() => {
        if (!item) return '';
        if (typeof item === 'string') {
          const parts = item.split('/');
          return parts[parts.length - 1];
        }
        return item.FileName || item.fileName || item.FileLeafRef || item.Name || item.name || '';
      })();

      if (fileName && initialNumericRepoId && context) {
        deleteRCAAttachment(initialNumericRepoId, fileName, context).catch((e) =>
          console.error('Failed to delete attachment from SharePoint', fileName, e)
        );
      }

      return { ...s, attachments: current.filter((_: any, i: number) => i !== index) };
    });

    setErrors(prev => {
      const next = { ...prev };
      delete next['attachments'];
      return next;
    });
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
    // clear action-specific error flags when user fills data
    setErrors(prev => {
      const next = { ...prev };
      delete next[`${actionKey}-actionPlan`];
      delete next[`${actionKey}-responsibility`];
      return next;
    });
  };

  // store responsibility as array of strings (email/login/name) so repository can accept string[] or string
  const handlePeoplePickerChange = (actionKey: string) => (items: any[]) => {
    const list = Array.isArray(items) ? items : [];
    const values = list
      .map((p: any) => {
        const id = (p && (p.id || p.Id || p.ID)) ?? '';
        const display = (p && (p.secondaryText || p.loginName || p.text || p.email || '')) || '';
        return `${String(id)}|${String(display)}`.trim();
      })
      .filter((v: any): v is string => typeof v === 'string' && v.length > 0);
    updateActionDetail(actionKey, 'responsibility', values);

    // clear responsibility error for the action
    setErrors(prev => {
      const next = { ...prev };
      delete next[`${actionKey}-responsibility`];
      return next;
    });
  };

  // validation routine: returns true if valid, sets errors otherwise
  const validate = (): boolean => {
    const nextErrors: Record<string, string> = {};

    if (!form.problemStatement || String(form.problemStatement).trim() === '') {
      nextErrors['problemStatement'] = 'Problem statement is required.';
    }
    if (!form.causeCategory || String(form.causeCategory).trim() === '') {
      nextErrors['causeCategory'] = 'Cause category is required.';
    }
    // require causes and root causes
    if (!form.causes || String(form.causes).trim() === '') {
      nextErrors['causes'] = 'Cause(s) is required.';
    }
    if (!form.rootCauses || String(form.rootCauses).trim() === '') {
      nextErrors['rootCauses'] = 'Root cause(s) is required.';
    }
    // require at least one action type
    if (!form.actionType || !Array.isArray(form.actionType) || form.actionType.length === 0) {
      nextErrors['actionType'] = 'Select at least one Type of Action.';
    } else {
      // per-action validations
      (form.actionType as string[]).forEach((act) => {
        const details = actionDetails[act] || {};
        if (!details.actionPlan || String(details.actionPlan).trim() === '') {
          nextErrors[`${act}-actionPlan`] = 'Action Plan is required for this action.';
        }
        const resp = details.responsibility;
        const hasResp = Array.isArray(resp) ? resp.length > 0 : !!resp;
        if (!hasResp) {
          nextErrors[`${act}-responsibility`] = 'Responsibility is required for this action.';
        }
      });
    }

    setErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  };

  // build item matching IRCAList fields expected by saveRCAItem
  const formatDate = (d: any) => {
    if (!d) return undefined;
    if (d instanceof Date && !isNaN(d.getTime())) return d.toISOString();
    return d;
  };
  // const fetchRCAItems = async () => {
  //   const genericServiceInstance: IGenericService = new GenericService(undefined, context);
  //   genericServiceInstance.init(undefined, context);
  //   const RCARepo: IRCARepository = new RCARepository(genericServiceInstance);
  //   RCARepo.setService(genericServiceInstance);
  //   const RAitems = await getRCAItems(true, context);
  //   setRCAItems(RAitems);
  // }
  const onSave = async () => {
    // run validation before preparing payload / saving
    if (!validate()) {
      // focus/scroll to first error could be added here
      return;
    }

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



    const item: any = {};
    // title / problem statement
    item.LinkTitle = form.problemStatement || '';
    //item.ProblemStatement = form.problemStatement || '';

    // top-level mappings
    item.CauseCategory = form.causeCategory || '';
    item.RCASource = form.source || '';
    item.RCAPriority = form.priority || '';
    item.RelatedMetric = form.relatedMetric || '';
    item.RelatedSubMetric = form.relatedSubMetric || '';
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
      let savedItemId: number | undefined = undefined;

      if (context) {
        if (numericRepoId && numericRepoId > 0) {
          try {
            await updateRCAItem(numericRepoId, item, context);
            console.log('RCA updated', numericRepoId);
            savedItemId = numericRepoId;
            const successMessage = numericRepoId && numericRepoId > 0 ? 'RCA updated successfully.' : 'RCA saved successfully.';
            showMessage(successMessage, 'success');
            // await fetchRCAItems();
          } catch (e: any) {
            console.error('Failed to update RCA item', e);
            showMessage('Failed to update RCA item. Please try again later.', 'error');
          }
        } else {
          try {
            const result = await saveRCAItem(item, context);
            console.log('RCA saved', result);
            // try to extract item id from repository result (common shapes)
            savedItemId =
              (result && (result.Id || result.ID || result.id)) ||
              (result && result.data && (result.data.Id || result.data.ID || result.data.id)) ||
              undefined;
            if (typeof savedItemId === 'string') savedItemId = Number(savedItemId);
            const successMessage = numericRepoId && numericRepoId > 0 ? 'RCA updated successfully.' : 'RCA saved successfully.';
            showMessage(successMessage, 'success');
            // await fetchRCAItems();
          } catch (e: any) {
            console.error('Failed to save RCA item', e);
            showMessage('Failed to save RCA item. Please try again later.', 'error');

          }
        }
      } else {
        console.warn('No WebPart context provided to save/update RCA item - skipping backend call.');
      }

      // upload new attachments (File objects)
      if (savedItemId && form.attachments && form.attachments.length > 0) {
        const filesToUpload = (form.attachments || []).filter((a: any) => a instanceof File) as File[];
        if (filesToUpload.length > 0) {
          for (const file of filesToUpload) {
            try {
              // savedItemId is checked above — assert as number for the repo API
              await uploadRCAAttachment(savedItemId as number, file, context);
              console.log('Uploaded attachment', file.name);
              window.location.reload();

            } catch (e: any) {
              console.error('Failed to upload attachment', file.name, e);
            }
          }
        }
      }





      // show modal and call parent callback (parent can also refetch onSubmit)

      if (onSubmit) onSubmit(payload);
      else console.log('RCA Form submit', payload);
    } catch (err: any) {
      console.error('Failed to save/update RCA item', err);
      showMessage('Failed to save RCA item. Please try again later.', 'error');
    }
  };

  // add reset handler
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
  // message captured when tenant people search fails
  //const [peopleSearchErrorMessage, setPeopleSearchErrorMessage] = useState<string>('');

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

  // listen for global errors/unhandled rejections coming from PeopleSearchService and mark fallback
  useEffect(() => {
    const handleWindowError = (ev: ErrorEvent) => {
      const msg = ev?.message || (ev?.error && ev.error.message) || '';
      if (msg && (msg.indexOf('PeopleSearchService::searchTenant') !== -1 || msg.indexOf('searchTenant') !== -1 || msg.indexOf('PeopleSearchService') !== -1)) {
        setPeoplePickerFailed(true);
        //setPeopleSearchErrorMessage('People search failed — use manual entry or retry.');
      }
    };

    const handleUnhandledRejection = (ev: any) => {
      const reason = ev?.reason?.message || String(ev?.reason || '');
      if (reason && (reason.indexOf('PeopleSearchService::searchTenant') !== -1 || reason.indexOf('searchTenant') !== -1 || reason.indexOf('PeopleSearchService') !== -1)) {
        setPeoplePickerFailed(true);
        // setPeopleSearchErrorMessage('People search failed — use manual entry or retry.');
      }
    };

    window.addEventListener('error', handleWindowError);
    window.addEventListener('unhandledrejection', handleUnhandledRejection);
    return () => {
      window.removeEventListener('error', handleWindowError);
      window.removeEventListener('unhandledrejection', handleUnhandledRejection);
    };
  }, []);

  useEffect(() => {
    if (context) {
      loadMetricsData().catch(() => {
        setMetricsData([]);
      });
    }
  }, [context]);
  useEffect(() => {
    if (context && form.relatedMetric !== "" || form.relatedMetric !== undefined) {

      loadSubMetricsData().catch(() => {
        setSubMetricsData([]);
        update('relatedSubMetric', undefined);
      });
    }
  }, [context, form.relatedMetric]);
  const loadMetricsData = async () => {
    const genericServiceInstance: IGenericService = new GenericService(undefined, context);
    genericServiceInstance.init(undefined, context);
    const MetricsMeasurementRepo: IProjectMetricsRepository = new MetricsRepository(genericServiceInstance);
    MetricsMeasurementRepo.setService(genericServiceInstance);

    let MetricValues = await getMetricsFromProjectMetrics(false, context)
    const mapped = MetricValues.map(m => ({ key: m.Metrics || '', text: m.Metrics || '' }));
    // keep only unique keys (preserve first occurrence)
    const unique: Array<{ key: string; text: string }> = [];
    mapped.forEach(m => {
      if (m.key && !unique.some(u => u.key === m.key)) unique.push(m);
    });

    const options = [{ key: 'None', text: 'None' }, ...unique];
    // setDropdownValueMetrics(ItemData["Metrics"]);
    setMetricsData(options);
  };
  const loadSubMetricsData = async () => {
    const genericServiceInstance: IGenericService = new GenericService(undefined, context);
    genericServiceInstance.init(undefined, context);
    const MetricsMeasurementRepo: IProjectMetricsRepository = new MetricsRepository(genericServiceInstance);
    MetricsMeasurementRepo.setService(genericServiceInstance);

    let MetricValues = await getSubMetricsFromProjectMetrics(false, context, form.relatedMetric)
    const mapped = MetricValues.map(m => ({ key: m.SubMetrics || '', text: m.SubMetrics || '' }));
    // keep only unique keys (preserve first occurrence)
    const unique: Array<{ key: string; text: string }> = [];
    mapped.forEach(m => {
      if (m.key && !unique.some(u => u.key === m.key)) unique.push(m);
    });

    const options = [{ key: 'None', text: 'None' }, ...unique];
    // setDropdownValueMetrics(ItemData["Metrics"]);
    setSubMetricsData(options);
  };

  // derive numeric repo id for existing item (used to fetch attachments)
  const initialRepoId = initialData ? (initialData.__repoId ?? initialData.__id ?? initialData.id ?? initialData.ID) : undefined;
  const initialNumericRepoId = initialRepoId ? Number(initialRepoId) : undefined;

  // track whether we already loaded existing attachments from server
  // const [existingAttachmentsLoaded, setExistingAttachmentsLoaded] = useState<boolean>(false);

  // helper: upload attachments to a SharePoint list item's AttachmentFiles collection
  //const RCA_LIST_TITLE = SubSiteListNames.RootCauseAnalysis; // adjust if your list has a different display title



  // helper to delete attachments by filename from a list item
  // const deleteAttachments = async (listTitle: string, itemId: number, fileNames: string[]) => {
  //   if (!sp || !listTitle || !itemId || !fileNames || fileNames.length === 0) return;
  //   const item = sp.web.lists.getByTitle(listTitle).items.getById(itemId);
  //   for (const fileName of fileNames) {
  //     try {
  //       await item.attachmentFiles.getByName(fileName).delete();
  //     } catch (e) {
  //       console.error('Error deleting attachment', fileName, e);
  //     }
  //   }
  // };

  // helper: normalize various attachment shapes into { FileName, ServerRelativeUrl } or keep File
  const normalizeAttachment = (a: any) => {
    if (!a) return null;
    if (a instanceof File) return a;
    if (typeof a === 'string') {
      const parts = a.split('/');
      return { FileName: parts[parts.length - 1] || a, ServerRelativeUrl: a };
    }
    // common SP shapes
    const fileName = a.FileName || a.fileName || a.FileLeafRef || a.Name || a.Title || a.name || (a.Url ? (a.Url.split('/').pop() || '') : '');
    const serverRelativeUrl = a.ServerRelativeUrl || a.Url || a.FileRef || (a.ServerRelativePath && a.ServerRelativePath.DecodedUrl) || '';
    return { FileName: fileName, ServerRelativeUrl: serverRelativeUrl };
  };

  // const sp = React.useMemo(() => resolveSPFromGenericService(context), [context]);
  const peoplePickerWebUrl = React.useMemo(() => {
    if (!context) return window.location.origin;
    return (context as any)?.pageContext?.web?.absoluteUrl ?? window.location.origin;
  }, [context]);

  return (
    <>
      {/* Message modal shown after successful save/update */}
      <MessageModal
        isOpen={showMessageModal}
        message={messageText}
        type={messageType}
        onDismiss={handleDismissMessage}
      />

      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <TextField
          label="Problem statement (Causal Analysis Trigger)"
          value={form.problemStatement}
          onChange={(_, v) => update('problemStatement', v)}
          errorMessage={errors['problemStatement'] || ''}
          required
        />

        <div style={{ display: 'flex', gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Dropdown
              label="Cause Category"
              options={causeCategoryOptions}
              selectedKey={form.causeCategory || undefined}
              onChange={(_, o) => {
                update('causeCategory', o?.key);
                // clear error if set
                setErrors(prev => { const next = { ...prev }; delete next['causeCategory']; return next; });
              }}
            />
            {errors['causeCategory'] && <div style={{ color: 'red', fontSize: 12, marginTop: 6 }}>{errors['causeCategory']}</div>}
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
          <Dropdown
            label="Related Metric (if any)"
            options={MetricsData}
            selectedKey={form.relatedMetric || undefined}
            onChange={(_, o) => {
              const key = o?.key;
              update('relatedMetric', key);
              // clear selected sub-metric when metric changes

              // refresh sub-metrics for newly selected metric

            }}
          />
          {((form.relatedMetric !== "" || form.relatedMetric !== undefined) && (form.relatedMetric !== "None" || SubMetricsData.length > 1)) && <Dropdown
            label="Related Sub Metric (if any)"
            options={SubMetricsData}
            selectedKey={form.relatedSubMetric || undefined}
            onChange={(_, o) => update('relatedSubMetric', o?.key)}
          />
          }
        </div>

        <TextField
          label="Cause(s)"
          value={form.causes}
          onChange={(_, v) => update('causes', v)}
          errorMessage={errors['causes'] || ''}
          required
        />

        <TextField
          label="Root Cause(s)"
          value={form.rootCauses}
          onChange={(_, v) => update('rootCauses', v)}
          errorMessage={errors['rootCauses'] || ''}
          required
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
          <label style={{ display: 'block', marginBottom: 6, fontSize: '14px', fontWeight: '600' }}>Type of Action</label>
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
                      <div style={{ fontSize: 16, fontWeight: 600 }}>{act}</div>
                    </div>
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
                      {!peoplePickerFailed && context ? (
                        <PeoplePickerErrorBoundary onError={() => setPeoplePickerFailed(true)}>
                          <PeoplePicker
                            context={context as any}
                            webAbsoluteUrl={peoplePickerWebUrl}
                            titleText=""
                            personSelectionLimit={5}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={300}
                            showtooltip
                            ensureUser
                            placeholder="Type a name or email..."
                            defaultSelectedUsers={
                              (() => {
                                const raw = actionDetails[act]?.responsibility;
                                if (!raw) return [];
                                if (Array.isArray(raw)) {
                                  return raw
                                    .map((r: any) => {
                                      if (typeof r === 'string') {
                                        // if stored as "id|value", use right side for picker display
                                        const parts = r.split('|');
                                        return (parts.length > 1 ? parts[1] : parts[0]).trim();
                                      }
                                      return (r.secondaryText || r.loginName || r.text || '').trim();
                                    })
                                    .filter(Boolean);
                                }
                                if (typeof raw === 'string' && raw.length) {
                                  return raw
                                    .split(/; ?/)
                                    .map((s: string) => {
                                      const parts = s.split('|');
                                      return (parts.length > 1 ? parts[1] : parts[0]).trim();
                                    })
                                    .filter(Boolean);
                                }
                                return [];
                              })()
                            }
                            onChange={handlePeoplePickerChange(act)}
                          />
                        </PeoplePickerErrorBoundary>
                      ) : (
                        <div style={{ marginBottom: 6, color: '#a19f9d', fontSize: 12 }}>
                          People search failed — use manual entry or retry later.
                        </div>
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
            onClick={() => {
              const next = !attachmentsOpen;
              setAttachmentsOpen(next);
              if (next && (!form.attachments || form.attachments.length === 0) && initialData?.attachments?.length) {
                const normalized = initialData.attachments
                  .map((a: any) => normalizeAttachment(a))
                  .filter(Boolean);
                setForm((s: any) => {
                  const cur = Array.isArray(s.attachments) ? [...s.attachments] : [];
                  const names = new Set(cur.map((c: any) => (c instanceof File ? c.name : (c.FileName || c.name || ''))));
                  normalized.forEach((n: any) => {
                    const fname = n instanceof File ? n.name : n.FileName || '';
                    if (fname && !names.has(fname)) {
                      cur.push(n);
                      names.add(fname);
                    }
                  });
                  return { ...s, attachments: cur };
                });
              }
            }}
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
                  {form.attachments.map((a: any, idx: number) => {
                    // File objects (new uploads)
                    if (a instanceof File) {
                      return (
                        <div key={`new-${idx}`} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 420 }}>{a.name}</span>
                          <DefaultButton text="Remove" onClick={() => removeAttachment(idx)} />
                        </div>
                      );
                    }

                    // existing attachment: could be string URL or object with Url/ServerRelativeUrl/FileName
                    const fileName = ((): string => {
                      if (typeof a === 'string') {
                        const parts = a.split('/');
                        return parts[parts.length - 1] || a;
                      }
                      return (a.FileName || a.fileName || a.Name || a.Title || a.FileLeafRef || a.name || '');
                    })();

                    const fileUrl = ((): string | undefined => {
                      if (typeof a === 'string') return a;
                      if (a.ServerRelativeUrl) return a.ServerRelativeUrl;
                      if (a.Url) return a.Url;
                      if (a.FileRef) return a.FileRef;
                      if (a.ServerRelativePath && a.ServerRelativePath.DecodedUrl) return a.ServerRelativePath.DecodedUrl;
                      return undefined;
                    })();

                    return (
                      <div key={`existing-${idx}`} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        {fileUrl ? (
                          <a href={fileUrl} target="_blank" rel="noreferrer" style={{ maxWidth: 420, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {fileName || fileUrl}
                          </a>
                        ) : (
                          <span style={{ maxWidth: 420, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{fileName || 'attachment'}</span>
                        )}
                        <DefaultButton text="Remove" onClick={() => removeAttachment(idx)} />
                      </div>
                    );
                  })}
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
    </>
  );
}