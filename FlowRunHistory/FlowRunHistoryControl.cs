using System;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using WinFormsLabel = System.Windows.Forms.Label;
using XrmToolBox.Extensibility;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PowerAutomateRunHistory
{
    public class FlowRunHistoryControl : PluginControlBase
    {
        // UI
        private ComboBox cboSolutions;
        private ComboBox cboFlows;
        private DateTimePicker dtStart;
        private DateTimePicker dtEnd;
        private TextBox txtEnvironmentId; // optional, to build run URL
        private Button btnLoad;
        private Button btnOpenSelected;
        private Button btnGetRunIO; // preview stub
        private DataGridView grid;
        private WinFormsLabel lblStatus;
        private Button btnExportExcel;
        //private TextBox txtFlowName;

        // Cache
        private Dictionary<Guid, string> solutions = new Dictionary<Guid, string>();
        private Dictionary<Guid, FlowInfo> flows = new Dictionary<Guid, FlowInfo>();

        private enum StatusType
        {
            Info,
            Success,
            Warning,
            Error
        }

        private Timer _statusFlashTimer;
        private int _statusFlashCounter;
        private Color _statusOriginalBackColor;

        private class FlowInfo
        {
            public Guid WorkflowId { get; set; }
            public Guid WorkflowIdUnique { get; set; }
            public string Name { get; set; }
            public string SolutionUniqueName { get; set; }
        }

        public FlowRunHistoryControl()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            SuspendLayout();
            AutoScroll = true;

            var panelTop = new Panel { Dock = DockStyle.Top, Height = 80 };
            var panelBtns = new Panel { Dock = DockStyle.Top, Height = 40 };
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoGenerateColumns = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            //lblStatus = new WinFormsLabel { Dock = DockStyle.Bottom, Height = 50, ForeColor = Color.DimGray, Text = "Ready", TextAlign = ContentAlignment.MiddleCenter };

            lblStatus = new WinFormsLabel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.RoyalBlue,
                BackColor = Color.AliceBlue,
                BorderStyle = BorderStyle.FixedSingle,
                Text = "Ready"
            };
            UpdateStatus("Ready", StatusType.Info, false);

            cboSolutions = new ComboBox { Left = 10, Top = 10, Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            cboFlows = new ComboBox { Left = 300, Top = 10, Width = 400, DropDownStyle = ComboBoxStyle.DropDownList };
            //var lblFlowName = new WinFormsLabel{ Left = 610, Top = 10, Width = 100, Text = "Flow Name (optional)" };
            //txtFlowName = new TextBox { Left = 720, Top = 10, Width = 260 };
            dtStart = new DateTimePicker { Left = 10, Top = 40, Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd HH:mm" };
            dtEnd = new DateTimePicker { Left = 200, Top = 40, Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd HH:mm" };
            var lblEnvId = new WinFormsLabel { Left = 390, Top = 44, Width = 80, Text = "Environment Id" };
            txtEnvironmentId = new TextBox { Left = 500, Top = 40, Width = 210, };

            btnLoad = new Button { Left = 610, Top = 10, Width = 140, Text = "Load Runs" };
            btnOpenSelected = new Button { Left = 760, Top = 10, Width = 160, Text = "Open Run in Portal" };
            btnGetRunIO = new Button { Left = 610, Top = 40, Width = 310, Text = "Get Run Inputs/Outputs (Preview)" };
            btnExportExcel = new Button { Left = 930, Top = 10, Width = 160, Text = "Export to Csv", Enabled = false };

            panelTop.Controls.AddRange(new Control[] { cboSolutions, cboFlows, dtStart, dtEnd, txtEnvironmentId, lblEnvId });
            panelBtns.Controls.AddRange(new Control[] { btnLoad, btnOpenSelected, btnExportExcel, btnGetRunIO });

            Controls.Add(grid);
            Controls.Add(panelBtns);
            Controls.Add(panelTop);
            Controls.Add(lblStatus);

            // Events
            Load += (_, __) => ExecuteMethod(LoadSolutions);
            cboSolutions.SelectedIndexChanged += (_, __) => ExecuteMethod(LoadFlowsForSelectedSolution);
            btnLoad.Click += (_, __) => ExecuteMethod(LoadRuns);
            btnOpenSelected.Click += (_, __) => ExecuteMethod(OpenSelectedRun);
            btnGetRunIO.Click += (_, __) => ExecuteMethod(GetRunIODetailsPreview);
            btnExportExcel.Click += (_, __) => ExecuteMethod(ExportRunsToCsv);

            // Grid columns
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "FlowName", HeaderText = "Flow", DataPropertyName = "FlowName" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "TriggerType", HeaderText = "Trigger", DataPropertyName = "TriggerType" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Solution", HeaderText = "Solution", DataPropertyName = "Solution" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "Status", DataPropertyName = "Status" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "StartTime", HeaderText = "Start", DataPropertyName = "StartTime" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DurationMs", HeaderText = "Duration (ms)", DataPropertyName = "DurationMs" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ErrorMessage", HeaderText = "Error", DataPropertyName = "ErrorMessage" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "RunId", HeaderText = "Run Id", DataPropertyName = "RunId" });
            grid.Columns.Add(new DataGridViewLinkColumn { Name = "RunUrl", HeaderText = "Run URL", DataPropertyName = "RunUrl", TrackVisitedState = true, Width = 75 });

            var copyCol = new DataGridViewButtonColumn
            {
                Name = "CopyRunUrl",
                HeaderText = "Copy URL",
                Text = "Copy",
                UseColumnTextForButtonValue = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                Width = 80,
                Resizable = DataGridViewTriState.False
            };

            grid.Columns.Add(copyCol);
            grid.Columns["CopyRunUrl"].DisplayIndex = grid.Columns.Count - 1;

            // Add right-click context menu for copying Run URL
            var ctxMenu = new ContextMenuStrip();
            var miCopyUrl = new ToolStripMenuItem("Copy Run URL");
            miCopyUrl.Click += (_, __) =>
            {
                if (grid.CurrentRow == null) return;
                var url = grid.CurrentRow.Cells["RunUrl"]?.Value?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(url))
                {
                    Clipboard.SetText(url);
                    UpdateStatus("Run URL copied to clipboard.", StatusType.Success, flash: false);
                }
                else
                {
                    MessageBox.Show("Run URL is empty. Set Environment Id and reload runs.", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            ctxMenu.Items.Add(miCopyUrl);
            grid.ContextMenuStrip = ctxMenu;

            grid.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var colName = grid.Columns[e.ColumnIndex].Name;

                if (colName == "RunUrl")
                {
                    var url = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(url))
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = url, UseShellExecute = true });
                }

                if (colName == "CopyRunUrl")
                {
                    // Prefer the 'RunUrl' cell for the same row
                    var urlCell = grid.Rows[e.RowIndex].Cells["RunUrl"];
                    var url = urlCell?.Value?.ToString() ?? "";

                    if (string.IsNullOrWhiteSpace(url))
                    {
                        MessageBox.Show("Run URL is empty. Set Environment Id and reload runs.", "Flow Run History",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    try
                    {
                        Clipboard.SetText(url);
                        UpdateStatus("Run URL copied to clipboard.", StatusType.Success, flash: false);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Unable to copy URL: {ex.Message}", "Flow Run History",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };

            ResumeLayout();
        }

        private void LoadSolutions()
        {
            UpdateStatus("Loading solutions...", StatusType.Info);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Query solutions...",
                Work = (w, a) =>
                {
                    var q = new QueryExpression("solution")
                    {
                        ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };
                    // Skip default solution (optional)
                    q.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);
                    var results = Service.RetrieveMultiple(q);

                    solutions.Clear();
                    foreach (var s in results.Entities)
                    {
                        var id = s.Id;
                        var friendly = s.GetAttributeValue<string>("friendlyname");
                        solutions[id] = friendly;
                    }

                    a.Result = results;
                },
                PostWorkCallBack = r =>
                {

                    if (r.Error != null)
                    {
                        UpdateStatus($"Error: {r.Error.Message}", StatusType.Error);
                        LogError(r.Error?.Message, "FlowRunHistory");
                        return;
                    }

                    cboSolutions.Items.Clear();
                    cboSolutions.Items.Add(new ComboItem<Guid>("---Select solution---", Guid.Empty));
                    foreach (var kv in solutions.OrderBy(k => k.Value))
                        cboSolutions.Items.Add(new ComboItem<Guid>(kv.Value, kv.Key));

                    if (cboSolutions.Items.Count > 0) cboSolutions.SelectedIndex = 0;
                    UpdateStatus($"Loaded {cboSolutions.Items.Count} solutions", StatusType.Success);
                },
                //ErrorHandler = ex => lblStatus.Text = $"Error: {ex.Message}"
            });
        }

        private void LoadFlowsForSelectedSolution()
        {
            var item = cboSolutions.SelectedItem as ComboItem<Guid>;
            if (item == null || (item?.Value == Guid.Empty))
            {
                cboFlows.Items.Clear(); return;
            }
            var solutionId = item.Value;

            UpdateStatus("Loading flows in solution...", StatusType.Info);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Query flows in solution...",
                Work = (w, a) =>
                {
                    // componenttype 29 = Workflow; category 5 = Modern Flow
                    // FetchXML join solutioncomponent -> workflow
                    var fetch = $@"
                                <fetch distinct='true'>
                                  <entity name='solutioncomponent'>
                                    <attribute name='objectid' />
                                    <attribute name='componenttype' />
                                    <filter>
                                      <condition attribute='solutionid' operator='eq' value='{solutionId}'/>
                                      <condition attribute='componenttype' operator='eq' value='29'/>
                                    </filter>
                                    <link-entity name='workflow' from='workflowid' to='objectid' alias='wf'>
                                      <attribute name='name'/>
                                      <attribute name='workflowid'/>
                                      <attribute name='workflowidunique'/>
                                      <attribute name='category'/>
                                      <filter>
                                        <condition attribute='category' operator='eq' value='5'/>
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
                    var results = Service.RetrieveMultiple(new FetchExpression(fetch));

                    flows.Clear();
                    foreach (var e in results.Entities)
                    {
                        var wf = (AliasedValue)e["wf.name"]; // Aliased values
                        var wfName = wf?.Value?.ToString() ?? "(no name)";
                        var wfId = ((Guid)((AliasedValue)e["wf.workflowid"]).Value);
                        var wfIdUnique = ((Guid)((AliasedValue)e["wf.workflowidunique"]).Value);

                        flows[wfId] = new FlowInfo
                        {
                            WorkflowId = wfId,
                            WorkflowIdUnique = wfIdUnique,
                            Name = wfName,
                            SolutionUniqueName = solutions[solutionId]
                        };
                    }
                    a.Result = flows.Values.ToList();
                },
                PostWorkCallBack = r =>
                {
                    if (r.Error != null)
                    {
                        UpdateStatus($"Error: {r.Error.Message}", StatusType.Error);
                        LogError(r.Error?.Message, "FlowRunHistory");
                        return;
                    }

                    cboFlows.Items.Clear();
                    foreach (var f in flows.Values.OrderBy(x => x.Name))
                        cboFlows.Items.Add(new ComboItem<Guid>($"{f.Name} ({f.WorkflowId})", f.WorkflowId));

                    if (cboFlows.Items.Count > 0) cboFlows.SelectedIndex = 0;
                    UpdateStatus($"Loaded {cboFlows.Items.Count} flows", StatusType.Success);
                },
            });
        }

        #region commented method
        //private void LoadRunsModified()
        //{
        //    // 1) Resolve target flow id: selected in dropdown or from typed name
        //    Guid wfId = Guid.Empty;
        //    FlowInfo flowInfo = null;

        //    // Try dropdown first
        //    var flowItem = cboFlows.SelectedItem as ComboItem<Guid>;
        //    if (flowItem != null && flows.TryGetValue(flowItem.Value, out var selectedInfo))
        //    {
        //        wfId = flowItem.Value;
        //        flowInfo = selectedInfo;
        //    }
        //    else
        //    {
        //        // If not selected, try to resolve by name typed in txtFlowName
        //        var typedName = (txtFlowName.Text ?? "").Trim();
        //        if (string.IsNullOrWhiteSpace(typedName))
        //        {
        //            // No selection and no name: proceed gracefully
        //            MessageBox.Show("Type a Flow Name or select a flow from the dropdown.", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            return;
        //        }

        //        // Query workflow table for Modern Flow (category = 5), optionally restricted to selected solution
        //        // Try exact match first, then case-insensitive match
        //        var q = new QueryExpression("workflow")
        //        {
        //            ColumnSet = new ColumnSet("workflowid", "workflowidunique", "name", "category"),
        //            Criteria = new FilterExpression(LogicalOperator.And)
        //        };
        //        q.Criteria.AddCondition("category", ConditionOperator.Equal, 5); // Modern Flow
        //        q.Criteria.AddCondition("name", ConditionOperator.Like, typedName);

        //        // Optional: restrict to flows in the selected solution (recommended)
        //        var solItem = cboSolutions.SelectedItem as ComboItem<Guid>;
        //        if (solItem != null)
        //        {
        //            // Join solutioncomponent -> workflow to ensure the flow belongs to the selected solution
        //            //var link = new LinkEntity("workflow", "solutioncomponent", "workflowid", "objectid", JoinOperator.Inner);
        //            //link.Columns = new ColumnSet("solutionid", "componenttype");
        //            //link.EntityAlias = "sc";
        //            //link.LinkCriteria = new FilterExpression(LogicalOperator.And);
        //            //link.LinkCriteria.AddCondition("solutionid", ConditionOperator.Equal, solItem.Value);
        //            //link.LinkCriteria.AddCondition("componenttype", ConditionOperator.Equal, 29); // 29 = Workflow
        //            //q.LinkEntities.Add(link);
        //        }

        //        EntityCollection ecExact = null;
        //        try { ecExact = Service.RetrieveMultiple(q); }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"Error resolving flow by name: {ex.Message}", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }

        //        Entity target = ecExact.Entities.FirstOrDefault();

        //        if (target == null)
        //        {
        //            // Try case-insensitive LIKE
        //            var qLike = new QueryExpression("workflow")
        //            {
        //                ColumnSet = new ColumnSet("workflowid", "workflowidunique", "name", "category"),
        //                Criteria = new FilterExpression(LogicalOperator.And)
        //            };
        //            qLike.Criteria.AddCondition("category", ConditionOperator.Equal, 5);
        //            qLike.Criteria.AddCondition("name", ConditionOperator.Like, typedName);

        //            if (solItem != null)
        //            {
        //                var link2 = new LinkEntity("workflow", "solutioncomponent", "workflowid", "objectid", JoinOperator.Inner);
        //                link2.Columns = new ColumnSet("solutionid", "componenttype");
        //                link2.EntityAlias = "sc";
        //                link2.LinkCriteria = new FilterExpression(LogicalOperator.And);
        //                link2.LinkCriteria.AddCondition("solutionid", ConditionOperator.Equal, solItem.Value);
        //                link2.LinkCriteria.AddCondition("componenttype", ConditionOperator.Equal, 29);
        //                qLike.LinkEntities.Add(link2);
        //            }

        //            EntityCollection ecLike = null;
        //            try { ecLike = Service.RetrieveMultiple(qLike); } catch { ecLike = new EntityCollection(); }
        //            target = ecLike.Entities.FirstOrDefault(e =>
        //                string.Equals(e.GetAttributeValue<string>("name"), typedName, StringComparison.OrdinalIgnoreCase));
        //            if (target == null) target = ecLike.Entities.FirstOrDefault(); // otherwise take first partial match
        //        }

        //        if (target == null)
        //        {
        //            MessageBox.Show($"No flow found with name '{typedName}' in the selected solution/date scope.", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            return;
        //        }

        //        wfId = target.GetAttributeValue<Guid>("workflowid");
        //        var wfUnique = target.Contains("workflowidunique") ? target.GetAttributeValue<Guid>("workflowidunique") : Guid.Empty;
        //        var wfName = target.GetAttributeValue<string>("name") ?? "(unnamed)";

        //        flowInfo = new FlowInfo
        //        {
        //            WorkflowId = wfId,
        //            WorkflowIdUnique = wfUnique != Guid.Empty ? wfUnique : wfId,
        //            Name = wfName,
        //            SolutionUniqueName = (cboSolutions.SelectedItem as ComboItem<Guid>)?.Text ?? "(solution)"
        //        };
        //        flows[wfId] = flowInfo; // cache it so downstream code can use flowInfo
        //    }

        //    // 2) Date range (UTC for query, displayed local)
        //    var start = dtStart.Value.ToUniversalTime();
        //    var end = dtEnd.Value.ToUniversalTime();

        //    lblStatus.Text = "Loading flow runs...";
        //    WorkAsync(new WorkAsyncInfo
        //    {
        //        Message = "Query flowrun records...",
        //        Work = (w, a) =>
        //        {
        //            // Wrap in try so e.Error is populated on failure
        //            try
        //            {
        //                var cols = new ColumnSet("name", "starttime", "endtime", "status", "errormessage", "durationinms", "workflowid", "ownerid", "triggertype");
        //                var qe = new QueryExpression("flowrun")
        //                {
        //                    ColumnSet = cols,
        //                    Criteria = new FilterExpression(LogicalOperator.And)
        //                };
        //                qe.Criteria.AddCondition("workflowid", ConditionOperator.Equal, wfId);
        //                qe.Criteria.AddCondition("starttime", ConditionOperator.OnOrAfter, start);
        //                qe.Criteria.AddCondition("endtime", ConditionOperator.OnOrBefore, end);
        //                qe.AddOrder("starttime", OrderType.Descending);

        //                var paging = new PagingInfo { Count = 500, PageNumber = 1 };
        //                qe.PageInfo = paging;

        //                var list = new List<dynamic>();
        //                EntityCollection ec;
        //                do
        //                {
        //                    ec = Service.RetrieveMultiple(qe);
        //                    foreach (var r in ec.Entities)
        //                    {
        //                        var runName = r.GetAttributeValue<string>("name");
        //                        var st = r.GetAttributeValue<DateTime?>("starttime");
        //                        var et = r.GetAttributeValue<DateTime?>("endtime");
        //                        var dur = r.GetAttributeValue<int?>("durationinms");
        //                        var status = r.GetAttributeValue<string>("status");
        //                        var error = r.GetAttributeValue<string>("errormessage");

        //                        var envId = txtEnvironmentId.Text?.Trim();
        //                        var runUrl = (!string.IsNullOrWhiteSpace(envId) && !string.IsNullOrWhiteSpace(runName))
        //                            ? $"https://make.powerautomate.com/environments/{envId}/flows/{flowInfo.WorkflowIdUnique}/runs/{runName}"
        //                            : "";

        //                        list.Add(new
        //                        {
        //                            Solution = flowInfo.SolutionUniqueName,
        //                            FlowName = flowInfo.Name,
        //                            Status = status,
        //                            StartTime = st?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
        //                            EndTime = et?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
        //                            DurationMs = dur?.ToString(),
        //                            ErrorMessage = error,
        //                            RunId = runName,
        //                            RunUrl = runUrl
        //                        });
        //                    }

        //                    if (ec.MoreRecords)
        //                    {
        //                        paging.PageNumber++;
        //                        paging.PagingCookie = ec.PagingCookie;
        //                        qe.PageInfo = paging;
        //                    }
        //                } while (ec.MoreRecords);

        //                a.Result = list;
        //            }
        //            catch (Exception ex)
        //            {
        //                // Re-throw to let WorkAsync marshal to e.Error
        //                throw;
        //            }
        //        },
        //        PostWorkCallBack = e =>
        //        {
        //            if (e.Error != null)
        //            {
        //                lblStatus.Text = $"Error: {e.Error.Message}";
        //                LogError(e.Error.Message, "FlowRunHistory");
        //                return;
        //            }

        //            var data = e.Result as List<dynamic> ?? new List<dynamic>();
        //            grid.DataSource = data;
        //            lblStatus.Text = $"Loaded {data.Count} runs";
        //        }
        //    });
        //} 
        #endregion

        private void LoadRuns()
        {
            var flowItem = cboFlows.SelectedItem as ComboItem<Guid>;
            if (flowItem == null)
            {
                //MessageBox.Show("Please select a flow.", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadRunsByDateOnly();
                return;
            }

            var wfId = flowItem.Value;
            var start = dtStart.Value.ToUniversalTime();
            var end = dtEnd.Value.ToUniversalTime();

            UpdateStatus("Loading flow runs...", StatusType.Info);

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Query flowrun records...",
                Work = (w, a) =>
                {
                    var cols = new ColumnSet("name", "starttime", "endtime", "status", "errormessage", "duration", "workflowid", "ownerid", "triggertype");
                    var qe = new QueryExpression("flowrun")
                    {
                        ColumnSet = cols,
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };
                    qe.Criteria.AddCondition("workflowid", ConditionOperator.Equal, wfId);
                    qe.Criteria.AddCondition("starttime", ConditionOperator.OnOrAfter, start);
                    qe.Criteria.AddCondition("endtime", ConditionOperator.OnOrBefore, end);
                    qe.AddOrder("starttime", OrderType.Descending);

                    var page = 1;
                    var paging = new PagingInfo { Count = 500, PageNumber = page };
                    qe.PageInfo = paging;

                    var list = new List<FlowRunRow>();
                    EntityCollection ec;
                    do
                    {
                        ec = Service.RetrieveMultiple(qe);
                        foreach (var r in ec.Entities)
                        {
                            var runName = r.GetAttributeValue<string>("name"); // logic app run id
                            var st = r.GetAttributeValue<DateTime?>("starttime");
                            var et = r.GetAttributeValue<DateTime?>("endtime");
                            var dur = r.GetAttributeValue<long?>("duration");
                            var status = r.GetAttributeValue<string>("status");
                            var error = r.GetAttributeValue<string>("errormessage");
                            var triggerType = r.GetAttributeValue<string>("triggertype");

                            var solutionName = flows[wfId].SolutionUniqueName;
                            var flowName = flows[wfId].Name;
                            var envId = txtEnvironmentId.Text?.Trim();
                            var flowGuidForUrl = flows[wfId].WorkflowIdUnique; // use unique id

                            // If Environment Id provided, build portal URL
                            var runUrl = (!string.IsNullOrWhiteSpace(envId) && !string.IsNullOrWhiteSpace(runName))
                                ? $"https://make.powerautomate.com/environments/{envId}/flows/{flowGuidForUrl}/runs/{runName}"
                                : "";

                            list.Add(new FlowRunRow
                            {
                                Solution = solutionName,
                                FlowName = flowName,
                                Status = status,
                                StartTime = st?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                                //EndTime = et?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                                DurationMs = dur?.ToString(),
                                ErrorMessage = error,
                                RunId = runName,
                                RunUrl = runUrl,
                                TriggerType = triggerType
                            });
                        }

                        if (ec.MoreRecords)
                        {
                            paging.PageNumber++;
                            paging.PagingCookie = ec.PagingCookie;
                            qe.PageInfo = paging;
                        }
                    } while (ec.MoreRecords);

                    a.Result = list;
                },
                PostWorkCallBack = r =>
                {

                    if (r.Error != null)
                    {
                        UpdateStatus($"Error: {r.Error.Message}", StatusType.Error);
                        LogError(r.Error?.Message, "FlowRunHistory");
                        return;
                    }

                    var data = r.Result as List<FlowRunRow> ?? new List<FlowRunRow>();
                    grid.DataSource = new BindingList<FlowRunRow>(data);
                    grid.Columns["CopyRunUrl"].DisplayIndex = grid.Columns.Count - 1;
                    btnExportExcel.Enabled = grid.Rows.Count > 0;
                    UpdateStatus($"Loaded {data.Count} runs", StatusType.Success);

                },
                //ErrorHandler = ex => lblStatus.Text = $"Error: {ex.Message}"
            });
        }

        private void LoadRunsByDateOnly()
        {
            var start = dtStart.Value.ToUniversalTime();
            var end = dtEnd.Value.ToUniversalTime();

            UpdateStatus("Loading flow runs for the selected date range...", StatusType.Info);

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Querying Dataverse flowrun records...",
                Work = (w, a) =>
                {
                    try
                    {
                        // 1) Page through flowrun by date range
                        var cols = new ColumnSet("name", "starttime", "endtime", "status", "errormessage", "duration", "workflowid", "ownerid", "triggertype");
                        var qe = new QueryExpression("flowrun")
                        {
                            ColumnSet = cols,
                            Criteria = new FilterExpression(LogicalOperator.And)
                        };
                        qe.Criteria.AddCondition("starttime", ConditionOperator.OnOrAfter, start);
                        qe.Criteria.AddCondition("endtime", ConditionOperator.OnOrBefore, end);
                        qe.AddOrder("starttime", OrderType.Descending);

                        var paging = new PagingInfo { Count = 500, PageNumber = 1 };
                        qe.PageInfo = paging;

                        var runs = new List<Entity>();
                        EntityCollection ec;
                        do
                        {
                            ec = Service.RetrieveMultiple(qe);
                            runs.AddRange(ec.Entities);

                            if (ec.MoreRecords)
                            {
                                paging.PageNumber++;
                                paging.PagingCookie = ec.PagingCookie;
                                qe.PageInfo = paging;
                            }
                        } while (ec.MoreRecords);

                        // 2) Collect distinct workflow ids from runs
                        var wfIds = runs
                            .Where(r => r.Contains("workflowid"))
                            .Select(r => r.GetAttributeValue<string>("workflowid"))
                            .Distinct()
                            .ToList();

                        // 3) Retrieve workflow (name, workflowidunique, category)
                        var wfNameById = new Dictionary<Guid, (string Name, Guid UniqueId)>();
                        if (wfIds.Count > 0)
                        {
                            // Batch via IN filter
                            var qWF = new QueryExpression("workflow")
                            {
                                ColumnSet = new ColumnSet("workflowid", "workflowidunique", "name", "category"),
                                Criteria = new FilterExpression(LogicalOperator.And)
                            };
                            qWF.Criteria.AddCondition("workflowid", ConditionOperator.In, wfIds.ToArray());
                            qWF.Criteria.AddCondition("category", ConditionOperator.Equal, 5); // Modern Flow
                            var wfEC = Service.RetrieveMultiple(qWF);
                            foreach (var wf in wfEC.Entities)
                            {
                                var id = wf.Id;
                                var uniqueId = wf.Contains("workflowidunique") ? wf.GetAttributeValue<Guid>("workflowidunique") : id;
                                var name = wf.GetAttributeValue<string>("name") ?? "(unnamed)";
                                wfNameById[id] = (name, uniqueId);
                            }
                        }

                        // 4) Retrieve solution names for those workflows (solutioncomponent join)
                        var solutionNameByWfId = new Dictionary<Guid, string>();
                        if (wfIds.Count > 0)
                        {
                            var fetch = $@"
                                        <fetch>
                                          <entity name='solutioncomponent'>
                                            <attribute name='solutionid' />
                                            <attribute name='objectid' />
                                            <filter>
                                              <condition attribute='componenttype' operator='eq' value='29' />
                                              <condition attribute='objectid' operator='in'>
                                                {string.Join("", wfIds.Select(id => $"<value>{id}</value>"))}
                                              </condition>
                                            </filter>
                                            <link-entity name='solution' from='solutionid' to='solutionid' alias='sol'>
                                              <attribute name='friendlyname' />
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                            var scEC = Service.RetrieveMultiple(new FetchExpression(fetch));
                            foreach (var sc in scEC.Entities)
                            {
                                var wfId = sc.GetAttributeValue<Guid>("objectid");
                                var friendly = (sc.Contains("sol.friendlyname")
                                    ? ((AliasedValue)sc["sol.friendlyname"]).Value?.ToString()
                                    : null) ?? "(solution)";
                                solutionNameByWfId[wfId] = friendly;
                            }
                        }

                        // 5) Build result rows
                        var envId = txtEnvironmentId.Text?.Trim();
                        var list = new List<FlowRunRow>();

                        foreach (var r in runs)
                        {
                            var runName = r.GetAttributeValue<string>("name");
                            var st = r.GetAttributeValue<DateTime?>("starttime");
                            var et = r.GetAttributeValue<DateTime?>("endtime");
                            var dur = r.GetAttributeValue<long?>("duration");
                            var status = r.GetAttributeValue<string>("status");
                            var error = r.GetAttributeValue<string>("errormessage");
                            var triggerType = r.GetAttributeValue<string>("triggertype");

                            var wfRef = r.GetAttributeValue<string>("workflowid");
                            //var wfId = wfRef ?? Guid.Empty;
                            Guid.TryParse(wfRef, out Guid wfId);

                            wfNameById.TryGetValue(wfId, out var wfTuple);
                            var flowName = wfTuple.Name ?? "(flow)";
                            var flowUniqueGuid = wfTuple.UniqueId != Guid.Empty ? wfTuple.UniqueId : wfId;

                            var solutionName = solutionNameByWfId.TryGetValue(wfId, out var sName) ? sName : "(solution)";

                            var runUrl = (!string.IsNullOrWhiteSpace(envId) && !string.IsNullOrWhiteSpace(runName))
                                ? $"https://make.powerautomate.com/environments/{envId}/flows/{flowUniqueGuid}/runs/{runName}"
                                : "";

                            list.Add(new FlowRunRow
                            {
                                Solution = solutionName,
                                FlowName = flowName,
                                Status = status,
                                StartTime = st?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                                //EndTime = et?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                                DurationMs = dur?.ToString(),
                                ErrorMessage = error,
                                RunId = runName,
                                RunUrl = runUrl,
                                TriggerType = triggerType
                            });
                        }

                        a.Result = list;
                    }
                    catch
                    {
                        throw; // Will be surfaced as e.Error in PostWorkCallBack
                    }
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        UpdateStatus($"Error: {e.Error.Message}", StatusType.Error);
                        LogError(e.Error.Message, "FlowRunHistory");
                        return;
                    }
                    var data = e.Result as List<FlowRunRow> ?? new List<FlowRunRow>();
                    grid.DataSource = new BindingList<FlowRunRow>(data);
                    grid.Columns["CopyRunUrl"].DisplayIndex = grid.Columns.Count - 1;
                    btnExportExcel.Enabled = grid.Rows.Count > 0;
                    UpdateStatus($"Loaded {data.Count} runs", StatusType.Success);
                }
            });
        }

        private void OpenSelectedRun()
        {
            if (grid.SelectedRows.Count == 0) { MessageBox.Show("Select a run row first."); return; }
            var url = grid.SelectedRows[0].Cells["RunUrl"].Value?.ToString();
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("Run URL is empty. Please set Environment Id.", "Flow Run History", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = url, UseShellExecute = true });
        }

        // Preview stub: If you later add MSAL and call the (preview) Flow Runs API to enrich I/O details
        // See: https://learn.microsoft.com/en-us/rest/api/power-platform/powerautomate/flow-runs/list-flow-runs
        private void GetRunIODetailsPreview()
        {
            MessageBox.Show(
                "This feature uses preview APIs to fetch trigger/action inputs/outputs and requires OAuth to https://api.powerplatform.com.\n" +
                "For now, use the Run URL column to open the portal and inspect I/O.\n\n" +
                "Docs: Flow Runs REST API (preview).",
                "Preview",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportRunsToCsv()
        {
            try
            {
                using (var sfd = new SaveFileDialog
                {
                    Title = "Export Flow Run History (CSV)",
                    Filter = "CSV (*.csv)|*.csv",
                    FileName = $"FlowRunHistory_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                    OverwritePrompt = true
                })
                {
                    if (sfd.ShowDialog() != DialogResult.OK) return;
                    var lines = new List<string>
                    {
                        "Flow,TriggerType,Status,Solution,Start,Duration(ms),Error,Run Id,Run URL"
                    };

                    foreach (DataGridViewRow r in grid.Rows)
                    {
                        if (r.IsNewRow) continue;
                        string cell(object o) => (o?.ToString() ?? "").Replace("\"", "\"\"");
                        var row = string.Join(",",
                            cell(r.Cells["FlowName"]?.Value),
                            cell(r.Cells["TriggerType"]?.Value),
                            cell(r.Cells["Solution"]?.Value),
                            cell(r.Cells["Status"]?.Value),
                            cell(r.Cells["StartTime"]?.Value),
                            cell(r.Cells["DurationMs"]?.Value),
                            cell(r.Cells["ErrorMessage"]?.Value),
                            cell(r.Cells["RunId"]?.Value),
                            $"\"{cell(r.Cells["RunUrl"]?.Value)}\""
                        );
                        lines.Add(row);
                    }

                    System.IO.File.WriteAllLines(sfd.FileName, lines);
                    UpdateStatus($"Exported {lines.Count - 1} rows to {sfd.FileName}", StatusType.Success);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"CSV export failed: {ex.Message}", "Export to CSV",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogError(ex.Message, "FlowRunHistory");
            }
        }

        private void FlashStatus(Color accent)
        {
            try
            {
                if (_statusFlashTimer != null)
                {
                    _statusFlashTimer.Stop();
                    _statusFlashTimer.Dispose();
                    _statusFlashTimer = null;
                }

                _statusOriginalBackColor = lblStatus.BackColor;
                _statusFlashCounter = 0;

                _statusFlashTimer = new Timer { Interval = 180 }; // 180 ms per blink
                _statusFlashTimer.Tick += (s, e) =>
                {
                    lblStatus.BackColor = (_statusFlashCounter % 2 == 0) ? accent : _statusOriginalBackColor;
                    _statusFlashCounter++;
                    if (_statusFlashCounter >= 6) // ~1 second of flashing
                    {
                        _statusFlashTimer.Stop();
                        lblStatus.BackColor = _statusOriginalBackColor;
                    }
                };
                _statusFlashTimer.Start();
            }
            catch
            {
                // If anything goes wrong, fail silently and keep status visible
            }
        }

        private void UpdateStatus(string message, StatusType type, bool flash = true)
        {
            // Friendly emoji/icon + colors per status kind
            string icon;
            Color fore, back, accent;

            switch (type)
            {
                case StatusType.Success:
                    icon = "✅";
                    fore = Color.ForestGreen;
                    back = Color.Honeydew;
                    accent = Color.LightGreen;
                    break;
                case StatusType.Warning:
                    icon = "⚠";
                    fore = Color.DarkOrange;
                    back = Color.LemonChiffon;
                    accent = Color.Khaki;
                    break;
                case StatusType.Error:
                    icon = "❌";
                    fore = Color.White;
                    back = Color.Firebrick; // strong background for errors
                    accent = Color.OrangeRed;
                    break;
                default: // Info
                    icon = "ℹ";
                    fore = Color.RoyalBlue;
                    back = Color.AliceBlue;
                    accent = Color.LightSkyBlue;
                    break;
            }

            // Apply visuals
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;
            lblStatus.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblStatus.ForeColor = fore;
            lblStatus.BackColor = back;
            lblStatus.BorderStyle = BorderStyle.FixedSingle;
            lblStatus.Text = $"{icon} {message}";

            if (flash && type == StatusType.Error) // flash for info/warning/error, not success
                FlashStatus(accent);
        }

        // Helper
        private class ComboItem<T>
        {
            public string Text { get; }
            public T Value { get; }
            public ComboItem(string text, T value) { Text = text; Value = value; }
            public override string ToString() => Text;
        }

        private class FlowRunRow
        {
            public string FlowName { get; set; }
            public string TriggerType { get; set; }
            public string Solution { get; set; }
            public string Status { get; set; }
            public string StartTime { get; set; }
            //public string EndTime { get; set; }
            public string DurationMs { get; set; }
            public string ErrorMessage { get; set; }
            public string RunId { get; set; }
            public string RunUrl { get; set; }
        }
    }
}
