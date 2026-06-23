using ExcelDragDropAddIn;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ExcelDragDropAddIn
{
    public class ExcelWindowHook : NativeWindow
    {
        public event Action<List<string>, POINT> ImagesDropped;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Win32.WM_DROPFILES)
            {
                // 시스템 윈도우 프로시저 안에서 발생하는 모든 예외를 차단하여 엑셀이 꺼지는 것을 방지합니다.
                try
                {
                    IntPtr hDrop = m.WParam;
                    uint fileCount = Win32.DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);
                    List<string> files = new List<string>();

                    for (uint i = 0; i < fileCount; i++)
                    {
                        // 중요: 260자 길이 제한을 해결하기 위해 버퍼의 크기를 최대 32,768자로 확장합니다.
                        StringBuilder sb = new StringBuilder(32768);
                        Win32.DragQueryFile(hDrop, i, sb, (uint)sb.Capacity);
                        files.Add(sb.ToString());
                    }

                    // 드롭된 위치의 클라이언트 좌표 취득
                    POINT clientPt;
                    Win32.DragQueryPoint(hDrop, out clientPt);

                    // 화면 전체 좌표(Screen Coordinates)로 변환
                    POINT screenPt = clientPt;
                    Win32.ClientToScreen(this.Handle, ref screenPt);

                    Win32.DragFinish(hDrop);

                    if (files.Count > 0 && ImagesDropped != null)
                    {
                        ImagesDropped(files, screenPt);
                    }
                }
                catch (Exception ex)
                {
                    // 예외가 발생할 경우, 조용히 로그를 출력하고 무시하여 엑셀의 실행 상태를 유지합니다.
                    System.Diagnostics.Debug.WriteLine("DragDrop Hook Error: " + ex.Message);
                }

                // 0을 반환하여 메시지가 정상 처리되었음을 윈도우 시스템에 알림
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }
    }
}