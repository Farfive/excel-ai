import { useState, useCallback } from 'react';
import { WorkbookInfo } from '../types';
import { uploadWorkbook } from '../services/api';

interface UseWorkbookReturn {
  workbookInfo: WorkbookInfo | null;
  isUploading: boolean;
  uploadError: string | null;
  fileName: string | null;
  upload: (file: File) => Promise<void>;
  reset: () => void;
}

export function useWorkbook(): UseWorkbookReturn {
  const [workbookInfo, setWorkbookInfo] = useState<WorkbookInfo | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);

  const upload = useCallback(async (file: File) => {
    setIsUploading(true);
    setUploadError(null);
    setFileName(file.name);
    try {
      const info = await uploadWorkbook(file);
      setWorkbookInfo(info);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : 'Upload failed';
      setUploadError(msg);
      setFileName(null);
    } finally {
      setIsUploading(false);
    }
  }, []);

  const reset = useCallback(() => {
    setWorkbookInfo(null);
    setUploadError(null);
    setFileName(null);
  }, []);

  return { workbookInfo, isUploading, uploadError, fileName, upload, reset };
}
