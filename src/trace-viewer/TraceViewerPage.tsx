import React, { useEffect, useState, useCallback } from 'react';
import PluginTraceLogViewer, { PluginTraceLogData } from '../content/components/PluginTraceLogViewer';
import './TraceViewerPage.css';

const STORAGE_KEY = 'd365_trace_log_popout_data';
const STORAGE_TAB_KEY = 'd365_trace_log_source_tab';

const TraceViewerPage: React.FC = () => {
  const [data, setData] = useState<PluginTraceLogData | null>(null);
  const [loading, setLoading] = useState(true);
  const [sourceTabId, setSourceTabId] = useState<number | null>(null);

  const loadData = useCallback(() => {
    chrome.storage.local.get([STORAGE_KEY, STORAGE_TAB_KEY], (result) => {
      const stored = result[STORAGE_KEY];
      if (stored) {
        setData(stored as PluginTraceLogData);
      } else {
        setData({ logs: [], error: 'No trace log data found. Open the trace log viewer from the D365 toolbar first.' });
      }
      if (result[STORAGE_TAB_KEY]) {
        setSourceTabId(result[STORAGE_TAB_KEY]);
      }
      setLoading(false);
    });
  }, []);

  useEffect(() => {
    loadData();

    // Listen for data pushes from the content script
    const listener = (message: any) => {
      if (message.type === 'TRACE_LOG_DATA_UPDATE' && message.data) {
        setData(message.data as PluginTraceLogData);
      }
    };
    chrome.runtime.onMessage.addListener(listener);
    return () => chrome.runtime.onMessage.removeListener(listener);
  }, [loadData]);

  const handleRefresh = useCallback(() => {
    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { type: 'TRACE_LOG_REFRESH_REQUEST' });
    }
  }, [sourceTabId]);

  const handleClear = useCallback(() => {
    setData({ logs: [], moreRecords: false });
  }, []);

  const handleClose = useCallback(() => {
    window.close();
  }, []);

  if (loading) {
    return (
      <div className="trace-viewer-page-loading">
        <div className="trace-viewer-page-spinner" />
        <p>Loading trace logs...</p>
      </div>
    );
  }

  return (
    <div className="trace-viewer-page">
      <PluginTraceLogViewer
        data={data}
        onClose={handleClose}
        onRefresh={handleRefresh}
        onClear={handleClear}
        popout
      />
    </div>
  );
};

export default TraceViewerPage;
