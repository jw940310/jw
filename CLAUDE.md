# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

PCRS_Standard_Ver1 is the HMI/control application for an automated **steel-profile plasma cutting robot system** (Hanbit ENG / NAR). It drives a Hyundai Hi5 robot, AJINEXTEK (AXL) and Fastech motion axes (conveyor, discharge, T-axis), and plasma cutting hardware to cut steel profiles (angle, channel, beam, pipe, flat bar, square tube). It is a single-window Windows Forms desktop app, C# on **.NET Framework 4.7.2** (`WinExe`, AnyCPU). UI and comments are in Korean; date markers in comments use `[yy.mm.dd]`.

## Build & run

- **Build:** Open `PCRS_Standard_Ver1.sln` in Visual Studio 2022 (≥17.9) and build, or `msbuild PCRS_Standard_Ver1.sln /p:Configuration=Debug`. This is a legacy non-SDK `.csproj`; use **MSBuild/Visual Studio, not `dotnet build`**.
- The solution also contains `PCRS_Standard` (a `.vdproj` Visual Studio Installer project). Building it requires the "Microsoft Visual Studio Installer Projects" extension; it is not needed for normal development.
- There is **no test project, no linter, and no CI** in this repo.
- **Running requires native DLLs + data files in the output dir** (`bin/Debug/`): `AXL.dll`, `HRRpcRg.dll` (+ `HRRpc*.dll`), `EziMOTIONPlusE.dll`/`EzBasicAxl.dll`, `Guna.UI2.dll`, `MySql.Data.dll`, plus `NAR_PCRS.mdb`, `NAR_PCRS.ini`, `NAR_PCRS.mot`. Without the physical hardware the app still launches and asks whether to continue in no-hardware mode (`Define.GWorkHWNone`).

## External dependencies (mostly P/Invoke)

| File(s) | Wraps | Notes |
|---|---|---|
| `AXA/AXD/AXHS/AXL/AXM.cs` | `AXL.dll` (AJINEXTEK motion library) | Vendor-generated P/Invoke headers; classes `CAXA/CAXD/CAXHS/CAXL/CAXM` live **outside any namespace**. Comments are mojibake (corrupted encoding) — **do not edit/reformat these**. |
| `Hi5.cs` | `HRRpcRg.dll` (Hyundai Hi5 robot RPC) | Namespace is `PCRS_Extension_Ver1`, not the app namespace. |
| `Fastech_DEFINE.cs`, `Fastech_LIB.cs` | `EziMOTIONPlusE.dll` (Fastech Ezi-MOTION) | Servo drives. |
| `MySql.Data.dll` | MySQL | Optional MES integration only (see below). |
| `Microsoft.Office.Interop.Excel` | Excel | Report/statistics export. |

Note several `HintPath`s point outside the repo (e.g. `..\..\KIRO_KebbiTest\packages\...` for Guna.UI2, and an absolute MySQL Connector path); the runnable copies live in `bin/Debug/`.

## Architecture — the big picture

**Two global anchors hold the whole app together; almost every class reads/writes them:**

1. **`Define` (`Define.cs`)** — a static "god class" of global state. All shared config and data live here as `static` fields, most prefixed `G` (e.g. `GMainSelect`, `GDBPath`, `GMain_IP`/`GRobot_IP`, the material code tables `GAngleCode`/`GChannelCode`/…, robot/axis flags, background `Thread`s). When you need cross-cutting state, it is almost always in `Define`.
2. **`Program.f_Main` (`Program.cs`)** — a global static reference to the single `frmMain` instance. Non-form classes manipulate the UI directly via `Program.f_Main.<control>` (e.g. `Program.f_Main.listBox_WorkStatus`, `Program.f_Main.dgvWorkPartList`). There is no MVVM/MVP layer.

Most behavior lives in **large static helper classes** (`func_*`, `Load_DB`, `Robot_JobExecute`) that operate on `Define` + `Program.f_Main`, rather than in the forms themselves.

**Startup flow** (`frmMain` ctor → `frmMain_Load`):
`Load_DB.DB_CodeLoad()` fills `Define` material-code arrays from the `.mdb` → single-instance check → DB path from registry (`DBPath`, else file dialog) → `func_VariousOP.ReadSetup()` reads `NAR_PCRS.ini` → `func_AXL.InitLibrary()` opens AXL board → load `NAR_PCRS.mot` motion params → servo-ON conveyor/discharge/T-axis → start `Define.Now_Thread` / `Define.Safety_Thread` background monitors.

**Data layer:** primary store is a **local Microsoft Access `.mdb`** accessed via **Jet OLEDB 4.0** (`Define.CONNECT_STR = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "`). `func_DataBase.cs` has generic OLEDB helpers; `Load_DB.cs` loads the material/thickness/speed code tables at startup. The Jet 4.0 provider is **32-bit only** — the process must run as 32-bit. The `.mdb` path is persisted in the Windows Registry. MES integration (`func_SeqAutoRun.MES_Out*`, `frmMain.MES_Out_Status`) writes to a separate **MySQL** DB and is gated by `Define.MES_OUT_USE` from the ini.

**Forms** (`frm*.cs` + `.Designer.cs`): `frmMain` (main HMI), `frmManual` (manual axis/robot control), `frmSystemSetting` (settings), `frmDrawManage` (도면/drawing management), `frmCenterCorrection` (기준점 보정/center calibration), `frmSelectPart`(`_Multi`), `frmWorkout`/`frmYield` (production & yield stats), `frmStatisticChoice`, plus dialogs `DlgSplash`, `DlgAutoMessage`. The `.Designer.cs`/`.resx` files are auto-generated — edit via the VS designer, not by hand.

**Sequence/automation logic** (`func_Seq*`): `func_SeqHome` (homing), `func_SeqAutoRun` (the main automatic-run state machine — material length calc, robot JOB dispatch, MES) , `func_SeqTEnd` (tail-end sequence).

**Cut-path / JOB generation** (`CalcPos/`): per-material geometry → robot cutting JOB position calculations. One file per profile type — `Angle.cs`, `Channel.cs`, `Flange.cs` (H-beam), `FlatBar.cs`, `Pipe.cs`, `Tube.cs` (square tube), `FreeMacro.cs` — plus `MakeJOB.cs` (assembles the robot JOB, partly via an external JOB-generation DLL through `Define.GDLL_J_Val`), `JOB_File.cs`, `Define_R.cs`. `func_DrawCode.cs` generates cut codes; `func_Nesting.cs` does block/part nesting and the work part list.

**Material type is selected globally** via `Define.GMainSelect`, used in `switch` blocks throughout (`MakeJOB`, `func_Seq*`, `CalcPos`):
`0` 앵글(angle) · `1` 부등변앵글(unequal-leg angle) · `2` 쟌넬(channel) · `3` I-beam · `4` H-beam/플렌지(flange) · `5` 파이프(pipe) · `6` 평철(flat bar) · 사각관(square tube).

## Conventions & gotchas

- Error handling is overwhelmingly `try/catch` ending in `MessageBox.Show(ex.Message, …)`; follow that pattern for consistency rather than introducing logging frameworks.
- Korean comments/strings are normal and expected — keep them. Use the existing `G`-prefixed naming when adding global state to `Define`.
- The runtime config keys live in `NAR_PCRS.ini` and are parsed by name in `func_VariousOP.ReadSetup()` (e.g. `Main_IP`, `Robot_IP`, `MES_OutUse`); add new settings there alongside the existing `case` entries and a matching `Define` field.
