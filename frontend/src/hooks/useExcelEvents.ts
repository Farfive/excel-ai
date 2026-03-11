import { useEffect, useRef } from 'react';
import { sendDeltaUpdate } from '../services/api';

/* global Office, Excel */
declare const Excel: any;
declare const Office: any;

export function useExcelEvents(workbookUuid: string | null): void {
  const bufferRef = useRef<string[]>([]);
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (!workbookUuid) return;
    if (typeof Office === 'undefined') return;

    let handler: any = null;

    Office.onReady(() => {
      Excel.run(async (ctx: any) => {
        handler = ctx.workbook.worksheets.onChanged.add((args: any) => {
          const addr = args.address;
          bufferRef.current.push(addr);

          if (timerRef.current) clearTimeout(timerRef.current);
          timerRef.current = setTimeout(async () => {
            const cells = [...bufferRef.current];
            bufferRef.current = [];
            if (cells.length > 0 && workbookUuid) {
              try {
                await sendDeltaUpdate(workbookUuid, cells);
              } catch (e) {
                console.error('Delta update failed:', e);
              }
            }
          }, 2000);
        });
        await ctx.sync();
      }).catch((e) => console.error('useExcelEvents setup error:', e));
    });

    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
      if (handler) {
        Excel.run(async (ctx) => {
          handler?.remove();
          await ctx.sync();
        }).catch(() => {});
      }
    };
  }, [workbookUuid]);
}
