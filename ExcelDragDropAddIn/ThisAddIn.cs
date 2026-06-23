using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelDragDropAddIn
{
    public partial class ThisAddIn
    {
        private ExcelWindowHook _hook;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // [추가된 코드] 일반 권한 탐색기에서 보내는 드래그 앤 드롭 메시지를 허용 처리 (보안 정책 우회)
            Win32.ChangeWindowMessageFilter(Win32.WM_DROPFILES, Win32.MSGFLT_ADD);
            Win32.ChangeWindowMessageFilter(Win32.WM_COPYGLOBALDATA, Win32.MSGFLT_ADD);

            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            this.Application.SheetActivate += Application_SheetActivate;

            // 실행 시 최초 활성 창 후킹
            HookActiveWindow();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnhookWindow();
        }

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            HookActiveWindow();
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            HookActiveWindow();
        }

        private void Application_SheetActivate(object Sh)
        {
            HookActiveWindow();
        }

        /// <summary>
        /// 활성화된 Excel 창 내부의 EXCEL7 핸들을 찾아 메시지 후킹을 설정합니다.
        /// </summary>
        private void HookActiveWindow()
        {
            try
            {
                IntPtr excelMainHwnd = (IntPtr)this.Application.Hwnd;
                IntPtr excel7Hwnd = FindExcel7(excelMainHwnd);

                if (excel7Hwnd != IntPtr.Zero)
                {
                    if (_hook == null)
                    {
                        _hook = new ExcelWindowHook();
                        _hook.ImagesDropped += OnImagesDropped;
                    }

                    if (_hook.Handle != excel7Hwnd)
                    {
                        _hook.ReleaseHandle();
                        _hook.AssignHandle(excel7Hwnd);
                        Win32.DragAcceptFiles(excel7Hwnd, true);
                    }
                }
            }
            catch (Exception)
            {
                // 필요시 내부 로그 작성
            }
        }

        private void UnhookWindow()
        {
            if (_hook != null)
            {
                _hook.ImagesDropped -= OnImagesDropped;
                _hook.ReleaseHandle();
                _hook = null;
            }
        }

        /// <summary>
        /// 엑셀 메인 윈도우 하위에서 시트 영역 격자창(EXCEL7 클래스)을 재귀적으로 탐색합니다.
        /// </summary>
        private IntPtr FindExcel7(IntPtr parent)
        {
            IntPtr hwnd = Win32.FindWindowEx(parent, IntPtr.Zero, "EXCEL7", null);
            if (hwnd != IntPtr.Zero) return hwnd;

            IntPtr child = IntPtr.Zero;
            while ((child = Win32.FindWindowEx(parent, child, null, null)) != IntPtr.Zero)
            {
                hwnd = FindExcel7(child);
                if (hwnd != IntPtr.Zero) return hwnd;
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// 드롭된 이미지들을 셀에 배치하고 크기를 조정합니다.
        /// </summary>
        private void OnImagesDropped(List<string> filePaths, POINT screenPt)
        {
            try
            {
                Excel.Window activeWindow = this.Application.ActiveWindow;
                if (activeWindow == null) return;

                // 드롭된 마우스 스크린 좌표 기준의 셀 확인
                object targetObj = activeWindow.RangeFromPoint(screenPt.X, screenPt.Y);
                Excel.Range targetCell = targetObj as Excel.Range;

                // 마우스 포인터 위치에 셀을 감지하지 못한 경우 현재 선택된 ActiveCell 사용
                if (targetCell == null)
                {
                    targetCell = this.Application.ActiveCell;
                }

                if (targetCell == null) return;

                Excel.Worksheet sheet = targetCell.Worksheet;
                int rowOffset = 0;

                // 1단계: 윈도우 자연어 정렬 규칙 적용
                filePaths.Sort(Win32.StrCmpLogicalW);

                // 2단계: 리본의 '이름 역순 정렬' 체크 여부에 따라 역순 정렬 적용
                if (Globals.Ribbons.DragDropRibbon.chkReverseSort.Checked)
                {
                    filePaths.Reverse();
                }

                foreach (string path in filePaths)
                {
                    // [안정성 보완 추가] 드롭된 대상이 실제 디렉토리에 존재하는 '파일'인지 사전 검증합니다.
                    if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
                    {
                        continue;
                    }

                    string ext = Path.GetExtension(path).ToLower();
                    if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".webp")
                    {
                        // 다중 파일 드롭 시 아래 행으로 순차 정렬 배치
                        Excel.Range cell = targetCell.Offset[rowOffset, 0];

                        float left = (float)(double)cell.Left;
                        float top = (float)(double)cell.Top;
                        float width = (float)(double)cell.Width;
                        float height = (float)(double)cell.Height;

                        // 이미지 삽입 및 문서 내에 물리적인 파일로 포함 저장 설정
                        Excel.Shape shape = sheet.Shapes.AddPicture(
                            path,
                            Office.MsoTriState.msoFalse, // LinkToFile
                            Office.MsoTriState.msoTrue,  // SaveWithDocument
                            left,
                            top,
                            width,
                            height
                        );

                        // 셀과 함께 이동 및 크기 조정 설정 (Placement = xlMoveAndSize)
                        shape.Placement = Excel.XlPlacement.xlMoveAndSize;

                        rowOffset++;
                    }
                }
            }
            catch (Exception ex)
            {
                // UI 스레드 상의 예외 알림
                MessageBox.Show("이미지 삽입 중 오류가 발생했습니다:\n" + ex.Message, "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #region VSTO에서 생성된 코드

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}