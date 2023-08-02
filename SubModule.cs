using TaleWorlds.CampaignSystem.Conversation;
using TaleWorlds.Core;
using TaleWorlds.MountAndBlade;
using Microsoft.Office.Interop.Excel;
using TaleWorlds.CampaignSystem;
using System.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;


namespace TextFounder
{
    public class SubModule : MBSubModuleBase
    {
        protected override void OnSubModuleLoad()
        {
            base.OnSubModuleLoad();

        }

        protected override void OnSubModuleUnloaded()
        {
            base.OnSubModuleUnloaded();

        }

        protected override void OnBeforeInitialModuleScreenSetAsRoot()
        {
            base.OnBeforeInitialModuleScreenSetAsRoot();

        }

        public override void OnAfterGameInitializationFinished(Game game, object initializerObject)
        {
            base.OnAfterGameInitializationFinished(game, initializerObject);
            SentenceFounder();
        }

        public void SentenceFounder()
        {
            List<string> list ;
            Relaction(out list);
            UpdateExcile(list.ToArray());
        }

        public void Relaction(out List<string> ProceedText)
        {
            ProceedText = new List<string>();
            string aimName = "_sentences";
            CampaignGameMode gameMode;
            Campaign campaign = Campaign.Current;
            ConversationManager conversationManager = campaign.ConversationManager;
            FieldInfo myFieldInfo;
            Type myType = typeof(ConversationManager);
            myFieldInfo = myType.GetField(aimName, BindingFlags.NonPublic|BindingFlags.Instance);
            if (myFieldInfo != null)
            {
            var sentenceList = myFieldInfo.GetValue(conversationManager);
             List<ConversationSentence> sentenceList2 = sentenceList as List<ConversationSentence>;
                foreach (ConversationSentence sentence in sentenceList2)
                {
                    if(sentence.Text==null)
                    {
                        continue;
                    }   
                    else
                    {
                        ProceedText.Add("Behaviors "+"无 "+sentence.Id+" "+sentence.Text);
                    }
                }
                Debug.WriteLine($"总共收录对话{ProceedText.Count}句");
            }
            else
            {
                Debug.WriteLine($"未发现目标变量");

            }

        }

        public static void CreatExcile()
        {
            Application xlApp = new Application();

            // 隐藏Excel窗口
         //   xlApp.Visible = false;
            // 禁用警告消息
           // xlApp.DisplayAlerts = false;

            Workbook xlWorkBook = xlApp.Workbooks.Add();
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];

            // 设置Excel表格名称
            xlWorkSheet.Name = "Text";

            // 写入数据类型
            xlWorkSheet.Cells[1, 1] = "类型";
            xlWorkSheet.Cells[1, 2] = "事件";
            xlWorkSheet.Cells[1, 3] = "ID";
            xlWorkSheet.Cells[1, 4] = "文本";


            // 保存Excel文件
            xlWorkBook.SaveAs("D:\\工作文件\\骑砍项目\\TextFounder\\Text.xlsx");

            // 关闭Excel文件
            xlWorkBook.Close();

        }

        public static void UpdateExcile(string[] str)
        {
            // 打开现有的Excel文件
            Application xlApp = new Application();
            try
            {
                Workbook workbook = xlApp.Workbooks.Open("D:\\工作文件\\骑砍项目\\TextFounder\\Text.xlsx");
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
                string[] strs = new string[4];
                foreach (string item in str)
                {
                    int row = 2;
                    strs = item.Split(' ');
                    for (int column = 0; column < strs.Length; column++)
                    {
                        worksheet.Cells[row, column] = strs[column];
                    }

                    row++;
                }


                // 保存Excel文件
                workbook.Save();

                // 关闭Excel文件
                workbook.Close();

            }
    catch(Exception ex)
            {
           //     CreatExcile();
            }

                xlApp.Quit();

            





        }
    }
}
