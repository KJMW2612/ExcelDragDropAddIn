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
        /// 드롭된 이미지들을 셀에 배치하고 크기를 조정합니다. (병합 셀 완벽 지원)
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
                    if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
                    {
                        continue;
                    }

                    string ext = Path.GetExtension(path).ToLower();
                    if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".webp")
                    {
                        // 다중 파일 드롭 시 순차 배치할 기준 셀 계산
                        Excel.Range baseCell = targetCell.Offset[rowOffset, 0];

                        // [병합 셀 대응]
                        // 해당 셀이 속해있는 전체 병합 영역(MergeArea)을 동적으로 가져옵니다.
                        // 일반 셀인 경우 자기 자신을 반환하므로 일반/병합 환경 모두 자연스럽게 호환됩니다.
                        Excel.Range targetArea = baseCell.MergeArea;

                        // 병합 영역 전체의 좌표와 크기를 계산합니다.
                        float left = (float)(double)targetArea.Left;
                        float top = (float)(double)targetArea.Top;
                        float width = (float)(double)targetArea.Width;
                        float height = (float)(double)targetArea.Height;

                        // 이미지 삽입 및 전체 병합 크기에 맞춤 리사이즈
                        Excel.Shape shape = sheet.Shapes.AddPicture(
                            path,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
                            left,
                            top,
                            width,
                            height
                        );

                        shape.Placement = Excel.XlPlacement.xlMoveAndSize;

                        // [오프셋 계산 개선]
                        // 일반 셀일 때는 1행씩 내려가고, 
                        // 병합 셀일 때는 해당 병합 영역이 차지하는 실제 행(Row)의 개수만큼 건너뛰어 다음 사진이 겹치지 않게 합니다.
                        rowOffset += targetArea.Rows.Count;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x800A03EC)
            {
                MessageBox.Show(
                    "이미지를 삽입할 수 없습니다. 아래 사항을 확인해 주세요:\n\n" +
                    "1. 현재 셀을 더블클릭하여 '글자 입력(편집)' 중인지 확인\n" +
                    "   (키보드 Esc 키를 누른 뒤 다시 시도해 주세요)\n\n" +
                    "2. 엑셀 상단에 '제한된 보기' 노란색 바가 떠 있는지 확인\n" +
                    "   (상단의 '편집 사용' 버튼을 클릭해 주세요)\n\n" +
                    "3. 현재 시트가 '시트 보호' 상태인지 확인\n" +
                    "   (검토 탭에서 시트 보호를 해제해 주세요)\n\n" +
                    "4. 압축을 풀지 않은 ZIP 파일 내부에서 직접 드래그했는지 확인\n" +
                    "   (로컬 일반 폴더로 파일을 복사한 뒤 드래그해 주세요)",
                    "이미지 삽입 제한됨",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("이미지 삽입 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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