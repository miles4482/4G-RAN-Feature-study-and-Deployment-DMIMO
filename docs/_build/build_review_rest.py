#!/usr/bin/env python3
"""Remaining chapters for the Feature Review Excel."""

from xlsx_style import DocBook

FIG = "/workspace/docs/figures"


def ch4_config(book):
    s = book.sheet("04-3 Intra-Cell Data Prep", "4.4.1.1  Intra-Cell Data Preparation", "Source pages 34–39. Set these MOs in this order: CoProcRes → DMIMOCluster → DMIMOClusterCell → CellAlgoSwitch → optional algos.")
    s.para("For SFN cell data preparation see the SFN feature document. The following tables are the extra data for a D-MIMO cluster.")
    s.h2("Table 4-2  Parameters in the CoProcRes MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Coordinate Process Resource ID", "CoProcRes.CoProcResId", "0 to 12"],
        ["BaseBand Equipment ID", "CoProcRes.BaseBandEqmId", "ID of baseband equipment where the co-processing board is configured. 255 or 0–23."],
        ["Work Mode", "CoProcRes.WorkMode", "Need not configure if only D-MIMO is supported. Select COORDINATING_PROCESSING if D-MIMO and other coordination features are supported."],
        ["Bundling Cluster Type", "CoProcRes.BundlingClusterType", "DMIMO if only D-MIMO. ADAPTIVE only if D-MIMO and other coordination features."],
    ])
    s.h2("Table 4-3  Parameters in the DMIMOCluster MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["D-MIMO Cluster ID", "DMIMOCluster.DMIMOClusterId", "0–65535. Even number if >4 pRRUs in LampSite. Unique mod 6 per RRU in multi-carrier. Unique within a BBU."],
        ["Coordinate Process Resource ID", "DMIMOCluster.CoProcResId", "255 or 0–12. 255 = no co-processing resources specified."],
        ["Beamforming Weight Normalized Mode", "DMIMOCluster.BfWeightNormalizeMode", "NEBF: optimal for paired BF UEs at D-MIMO cell edge. PEBF: optimal at center / medium distance. Set NEBF in commercial use."],
    ])
    s.h2("Table 4-4  Parameters in the DMIMOClusterCell MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["D-MIMO Cluster ID", "DMIMOClusterCell.DMIMOClusterId", "0–65535"],
        ["Cell ID", "DMIMOClusterCell.CellId", "0–255"],
        ["eNodeB ID", "DMIMOClusterCell.eNodeBId", "0–1048575"],
        ["Mobile Country Code", "DMIMOClusterCell.Mcc", "000–999"],
        ["Mobile Network Code", "DMIMOClusterCell.Mnc", "00–99 or 000–999"],
    ])
    s.note(
        "Only SFN cells and pRRU aggregation cells can be added to a D-MIMO cluster. Only one cell can be added to each intra-cell D-MIMO cluster. The same cell cannot be added to different clusters. Add by primary operator ID or secondary, not both. Recommend primary operator ID.",
        "NOTICE",
    )
    s.h2("Table 4-5  Parameters in the CellAlgoSwitch MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["BF algorithm switch", "CellAlgoSwitch.BfAlgoSwitch", "Select BfSwitch."],
        ["MUBF Algorithm Switch", "CellAlgoSwitch.MuBfAlgoSwitch", "Select MuBfSwitch."],
        ["D-MIMO Algorithm Switch", "CellAlgoSwitch.DMIMOAlgoSwitch", "Inter-RRU cal and JT are enabled when DmimoJTSwitch is selected."],
        ["Enhanced Channel Calibration Switch", "CellAlgoSwitch.EnhChnCalSwitch", "Recommend QUICK_CHN_CAL_SWITCH when D-MIMO is enabled."],
        ["Coordination Algorithm Switch", "CellAlgoSwitch.CoordinationAlgoSwitch", "Recommend MULTI_UE_COORDINATION_OPT_SW when D-MIMO is enabled."],
    ])
    s.h2("Table 4-6  CellPcAlgo")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["D-MIMO Srs Power Control Offset for Sinr Target", "CellPcAlgo.DMSrsPcSinrOffset", "Recommend 5 when D-MIMO enabled. Default 0 when disabled."],
    ])
    s.h2("Table 4-7  CellBf")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Wait Pairing Layer Number Threshold", "CellBf.WaitPairingLayerThd", "Recommend 30 when cell load is heavy and multi-layer pairing rate is high. Recommend 0 otherwise."],
    ])
    s.h2("Table 4-8  CellPdcchAlgo")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["PDCCH Beamforming Gain Offset", "CellPdcchAlgo.PdcchBfGainOffset", "Recommend 0 when D-MIMO is on AND PDCCHAggLvlAdaptStrage = STRATEGYBASEDONCOVERAGE. Default -127 when D-MIMO is off."],
    ])
    s.h2("Table 4-9  CellBfMimoParaCfg")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Transmission Mode Acceleration Switch", "CellBfMimoParaCfg.TmAccelerationSwitch", "Recommend INITIAL_ACCESS_TO_BF when D-MIMO is on. Default OFF when disabled."],
    ])
    s.h2("Table 4-10  CellDlschAlgo")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Downlink High Load SDMA Threshold Offset", "CellDlschAlgo.DlHighLoadSdmaThdOffset", "Recommend 10 if D-MIMO + quick entry into BF; 5 if only D-MIMO in micro-micro. Default 0 when both disabled in micro-micro. Not required in macro-macro or macro-micro."],
    ])
    s.h2("Parameters used for optimization — SRS resource allocation (LampSite)")
    s.para(
        "SRSCfg.SrsResOptSwitch increases the proportion of UEs with short-period SRS, improving BF and DL spectral efficiency. Configure only when ALL of: D-MIMO enabled; UBBP boards manually bound; SRSCfg.TddSrsCfgMode = ACCESS_ENHANCED; Cell.MultiRruCellMode = MPRU_AGGREGATION."
    )
    s.note(
        "Modifying SrsResOptSwitch is a HIGH-RISK operation. Requires UBBPe / UBBPd6 / UBBPd9. UBBPd4 and LBBPd do not support it — enabling on those boards prevents physical cell activation and reports ALM-29243 Cell Capability Degraded or ALM-29240 Cell Unavailable. DSP CELL Work Status: 'The BBP configuration does not support SRS resource optimization'.",
        "NOTICE",
    )
    s.footer_end()


def ch4_mml(book):
    s = book.sheet("04-4 Intra-Cell MML", "4.4.1.2–4.4.1.3  Intra-Cell MML and MAE-Deployment", "Source pages 39–41. Complete 4.3 Requirements first. Check 'Service Interrupted After Modification' and Caution in parameter reference.")
    s.para(
        "Example: enable D-MIMO for SFN cell LocalCellId=0. UBBP in slots 0–3; slot 3 is co-processing. LampSite: four sector equipment. Macro: three sector equipment bound to slots 0–2. Set up and activate the SFN cell first (see SFN)."
    )
    s.mml("//Specifying baseband equipment for the BBP and co-processing board",
"""ADD BASEBANDEQM: BASEBANDEQMID=0, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=0;
ADD BASEBANDEQM: BASEBANDEQMID=1, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=1;
ADD BASEBANDEQM: BASEBANDEQMID=2, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=2;
ADD BASEBANDEQM: BASEBANDEQMID=3, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=3;""")
    s.mml("//(Only for macro eNodeBs) Adding sector equipment and binding baseband equipment for physical cells in SFN",
"""ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=0, BaseBandEqmId=0;
ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=1, BaseBandEqmId=1;
ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=2, BaseBandEqmId=2;""")
    s.mml("//(Only for LampSite eNodeBs) Sector equipment groups",
"""ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=0;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=1;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=2;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=3;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=0, SectorEqmId=0;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=1, SectorEqmId=1;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=2, SectorEqmId=2;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=3, SectorEqmId=3;""")
    s.mml("//Cluster objects and recommended switches",
"""ADD COPROCRES: CoProcResId=0, BaseBandEqmId=3, BundlingClusterType=DMIMO;
ADD DMIMOCLUSTER: DMIMOClusterId=0, CoProcResId=0;
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=0, eNodeBId=0, Mcc="100", Mnc="01";
MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=5;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-1, MuBfAlgoSwitch=MuBfSwitch-1, DMIMOAlgoSwitch=DmimoJTSwitch-1, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-1;
MOD CELLALGOSWITCH: LocalCellId=0, CoordinationAlgoSwitch= MULTI_UE_COORDINATION_OPT_SW-1;
MOD CELLPDCCHALGO: LocalCellId=0, PdcchBfGainOffset=0;
MOD CELLBF: LocalCellId=0, WaitPairingLayerThd=30;""")
    s.h2("Deactivation command examples")
    s.para("The following provides only deactivation examples. Restore other parameters based on actual network conditions.")
    s.mml("Deactivation (intra-cell)",
"""MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=0;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-0, MuBfAlgoSwitch=MuBfSwitch-0, DMIMOAlgoSwitch=DmimoJTSwitch-0, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-0;
MOD CELLALGOSWITCH: LocalCellId=0, CoordinationAlgoSwitch= MULTI_UE_COORDINATION_OPT_SW-0;
RMV DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=0, eNodeBId=0, Mcc="100", Mnc="01";
RMV DMIMOCLUSTER: DMIMOClusterId=0;
MOD CELLPDCCHALGO: LocalCellId=0, PdcchBfGainOffset=-127;
MOD CELLBF: LocalCellId=0, WaitPairingLayerThd=0;""")
    s.h2("4.4.1.3  Using the MAE-Deployment")
    s.para("For detailed operations, see Feature Configuration Using the MAE-Deployment.")
    s.note("Due to differences between OSS versions, GUIs in the interactive video may differ from actual OSS GUIs and are for reference only.", "NOTE")
    s.footer_end()


def ch4_verify(book):
    s = book.sheet("04-5 Intra-Cell Verify", "4.4.2–4.4.4  Activation Observation, Monitoring, Possible Issues", "Source pages 41–49. Exact MAE GUI screenshots from the PDF.")
    s.h2("4.4.2  Activation Observation")
    s.step(1, "DSP CELL and LST CELL — cell status of the cluster",
           "The cells in the D-MIMO cluster are working properly if ALL of the following are met:",
           "LST CELL: Mode of Multi-RRU Cell = SFN or MPRU_AGGREGATION. DSP CELL: Work Status = Normal for all RRUs. DSP CELL: Cell instance state = Normal.")
    s.step(2, "DSP DMIMOCLUSTERCELL — cluster cell status",
           "The D-MIMO cluster is working properly if D-MIMO Cluster Cell Status = Normal.",
           "DSP DMIMOCLUSTERCELL")
    s.step(3, "DSP DMIMOCALIBRATION — channel calibration",
           "Channel calibration is successful if D-MIMO Calibration Result = Succeeded.",
           "DSP DMIMOCALIBRATION")
    s.h2("4.4.3  Network Monitoring")
    s.h3("Table 4-11  Counters — UE pairing for D-MIMO")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.ChMeas.DMIMO.1Layer.PRB", "Average number of PRBs that can be paired for D-MIMO. D-MIMO takes effect if this value is > 0."],
        ["L.ChMeas.DMIMO.2Layer.PairPRB", "Average number of PRBs paired for D-MIMO at layer 2"],
        ["L.ChMeas.DMIMO.3Layer.PairPRB", "Average number of PRBs paired at layer 3"],
        ["L.ChMeas.DMIMO.4Layer.PairPRB", "Average number of PRBs paired at layer 4"],
        ["L.ChMeas.DMIMO.5Layer.PairPRB", "Average number of PRBs paired at layer 5"],
        ["L.ChMeas.DMIMO.6Layer.PairPRB", "Average number of PRBs paired at layer 6"],
        ["L.ChMeas.DMIMO.7Layer.PairPRB", "Average number of PRBs paired at layer 7"],
        ["L.ChMeas.DMIMO.8Layer.PairPRB", "Average number of PRBs paired at layer 8"],
    ])
    s.h3("Table 4-12  Proportion and scope of D-MIMO UEs")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.ChMeas.DMIMO.JT.User.Avg", "Number of D-MIMO JT UEs in the downlink"],
        ["L.ChMeas.DMIMO.JTUser.RRU.Avg", "Average number of working RRUs for D-MIMO JT UEs in the downlink"],
    ])
    s.h3("CHR: PERIOD_CLUSTER_DMIMO_MR (default 15 min)")
    s.table([
        ["Field", "Description"],
        [">DmimoClusterID", "ID of the D-MIMO cluster to which the logical cells performing MR reporting belong"],
        [">ClusterAbnormallatency", "Invalidity duration of the D-MIMO cluster"],
        [">CalibrationAbnormallatency", "Channel calibration invalidity duration of the D-MIMO cluster"],
    ])
    s.h3("MAE-Access MIMO (Cell) monitoring")
    s.table([
        ["Monitoring Item", "Unit", "Description"],
        ["Number of Enable DMIMO Pairing RB", "Number", "RBs that can be paired for D-MIMO in a cell in a monitoring period (averaged over TTIs)."],
        ["Number of successful DMIMO Pairing TM7 RB(Num) with N Layers", "Number", "Successfully paired RBs in TM7 at layer N. N = {2..8}."],
        ["Number of successful DMIMO Pairing TM8 RB(Num) with N Layers", "Number", "Successfully paired RBs in TM8 at layer N. N = {2..8}."],
        ["Number of successful DMIMO Pairing TM9 RB(Num) with N Layers", "Number", "Successfully paired RBs in TM9 at layer N. N = {2..8}."],
    ])
    s.h3("MAE-Access MIMO (User) monitoring (Test Items = DMIMO)")
    s.table([
        ["Monitoring Item", "Unit", "Description"],
        ["Number of Enable DMIMO Pairing RB(Num)", "Number", "RBs that can be paired for a D-MIMO UE. Displayed only when Test Items = DMIMO."],
        ["Number of successful DMIMO pairing RB (Num) with N Layers", "Number", "RBs for UEs successfully paired at layer N, averaged as cell-level (all RRUs). N={2..8}. Range 0–1000."],
    ])
    s.h3("CEU performance counters")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.Thrp.bits.UL.BorderUE.JointReception", "Total UL PDCP-layer traffic jointly received for CEUs"],
        ["L.Thrp.bits.UL.SmallPkt.BorderUE.JointReception", "PDCP traffic for UL small packets jointly received for CEUs"],
        ["L.Thrp.Time.UL.RmvSmallPkt.BorderUE.JointReception", "UL duration except small packets jointly received for CEUs"],
        ["L.Thrp.Time.UL.BorderUE.JointReception", "Total duration of jointly receiving data from CEUs (UL PDCP)"],
        ["L.Thrp.bits.DL.BorderUE.JointTransmit", "Total DL PDCP traffic jointly transmitted for CEUs"],
        ["L.Thrp.bits.DL.LastTTI.BorderUE.JointTransmit", "DL PDCP jointly transmitted for CEUs in last TTI before buffer empty"],
        ["L.Thrp.Time.DL.RmvLastTTI.BorderUE.JointTransmit", "JT CEU duration except last TTI before DL buffer empty"],
        ["L.Thrp.Time.DL.BorderUE.JointTransmit", "Total duration of jointly transmitting to CEUs (DL PDCP)"],
        ["L.ChMeas.PRB.PUSCH.Avg.BorderUE.JointReception", "Average PUSCH PRBs used by CEUs for joint reception"],
        ["L.ChMeas.PRB.PDSCH.Avg.BorderUE.JointTransmit", "Average PDSCH PRBs used by CEUs for joint transmission"],
    ])
    s.h2("4.4.4  Possible Issues")
    s.para("D-MIMO cells automatically revert to SFN cells if there is a fault related to channel calibration, activation of a cell in a cluster, or RRUs.")
    s.h3("Channel calibration failure alarm on MAE-Access")
    s.para("Configure a user-defined performance threshold alarm. Counter L.CellSectorEqpt.UNA.Dur.Cali measures calibration-failure duration.")
    s.step(1, "Open Threshold Settings",
           "MAE-Access: Performance > Threshold Management > Threshold Settings.",
           "Exact GUI from the source document:")
    s.figure(f"{FIG}/mae_step1_threshold_settings.png", "Step 1–2  MAE-Access Threshold Settings — choose Measurement of CellSectorEQUIP Performance, then Add")
    s.step(2, "Choose measurement and Add",
           "Navigation tree: eNodeB > Measurement of CellSectorEQUIP Performance. Click Add. Add Threshold dialog opens.")
    s.step(3, "Select NEs (Object tab)",
           "Select some or all NEs in the Object tab.")
    s.figure(f"{FIG}/mae_step3_add_threshold_object.png", "Step 3  Add Threshold — Object tab (example LTE_49/All Objects)")
    s.step(4, "Basic tab — Name, Period, Activation time",
           "Example: Name = CAL Fail; Period = 15 minutes; activation 00:00–23:59.")
    s.figure(f"{FIG}/mae_step4_basic_tab.png", "Step 4  Add Threshold — Basic tab (Name CAL Fail, Period 15 minutes)")
    s.step(5, "Advanced tab — counter and formula",
           "Alarm is reported if duration ≥ Threshold + Offset. Cleared if duration falls below Threshold − Offset. Direction=Ascending. Function subset=Network / Measurements Related to Algorithm(LTE) / Cell Sector Algorithm Measurement. Counter=L.CellSectorEqpt.UNA.Dur.Cali. Threshold = Period×80% − Offset (example 15×80%−3 = 9). Offset = Period×20% (example 3).")
    s.figure(f"{FIG}/mae_step5_advanced_tab.png", "Step 5  Add Threshold — Advanced tab (Ascending, L.CellSectorEqpt.UNA.Dur.Cali, Threshold 9, Offset 3)")
    s.h3("Channel calibration-related troubleshooting (DSP DMIMOCALIBRATION)")
    s.table([
        ["D-MIMO Calibration Result", "Meaning", "Action"],
        ["Succeeded", "Calibration OK", "None"],
        ["Exception occurred in internal calibration", "Intra-RRU channel calibration failed", "DSP CELLCALIBRATION to identify faulty RRU; contact Huawei TAC"],
        ["Route search failed", "No routing path between RRUs; external interference", "Eliminate interference first"],
        ["Reciprocity calibration failed", "Route search OK but reciprocity failed thresholds", "Collect Calibration Signal CINR and RSSI; contact Huawei TAC"],
    ])
    s.h3("Faults related to activation (DSP DMIMOCLUSTERCELL status ≠ Normal)")
    s.table([
        ["Status", "Check"],
        ["Incorrect configurations or unavailable licenses", "(1) D-MIMO switch on (2) RRU count proper (3) cluster configured (4) license activated"],
        ["Limited hardware capacity", "Cell set up on UBBPd/UBBPe?"],
        ["Route application failures", "Routing bandwidth sufficient?"],
        ["Cell abnormal", "DSP CELL"],
        ["Insufficient co-processing resources", "Co-processing board vs 4.3.3 Hardware"],
        ["Co-processing resources abnormal", "Co-processing board vs 4.3.3 Hardware"],
        ["Channel calibration failures", "DSP DMIMOCALIBRATION"],
        ["Cluster ID conflicts", "Cluster ID vs 4.3.4 Networking restraints"],
        ["Clusters being established", "Wait until setup succeeds"],
    ])
    s.h3("Others")
    s.para("When DmimoJTSwitch is selected, consecutive block errors may arise during dual-stream beamforming, even leading to network disconnections. Workarounds: see Beamforming (TDD).")
    s.footer_end()


def ch5(book):
    s = book.sheet("05 Inter-eNodeB D-MIMO", "5  Inter-eNodeB D-MIMO (TRIAL)", "Source pages 50–56. Feature TDLEOFD-121501. Depends on downlink D-MIMO (chapter 4).")
    s.note(
        "TRIAL FEATURE. Contact Huawei, sign MoU, understand risk. Not for sale in current version. May be invalidated on upgrade if not licensed when converted to commercial.",
        "TRIAL",
    )
    s.h2("5.1  Principles")
    s.para(
        "Cells generated by RRUs connected to different BBUs but on the same Cloud BB network form a D-MIMO cluster. Uses inter-eNodeB inter-RRU channel calibration and coherent joint transmission (see chapter 3). This scenario allows inter-BBU RRU networking and requires full frequency synchronization across these BBUs. If RGPS is the clock source of the entire Cloud BB network, the BBU on which the RGPS device is configured is available ONLY for intra-eNodeB D-MIMO, not inter-eNodeB. This restriction does not apply when a GPS clock is the source."
    )
    s.figure(f"{FIG}/fig_5_1_inter_enodeb_networking.png", "Figure 5-1  Networking diagram of inter-eNodeB D-MIMO (master reference clock vs system clock path)")
    s.h3("Inter-BBU clock solution")
    s.bullets([
        "BBU 1 connects to the clock reference source (e.g. RGPS), locks it, and shares the clock with the USUs.",
        "The USUs provide a system clock source. If two levels of USUs: second-level USU is the root. If one level: first-level USU is the root.",
        "Hardware synchronization modules of other BBUs lock to the system clock from step 2.",
    ], numbered=True)
    s.h2("5.2  Network Analysis")
    s.para("Benefits: same as downlink D-MIMO (4.2.1). Impacts: same as 4.2.2 / 5.3.2.3 (none extra).")
    s.h2("5.3.1  Licenses")
    s.table([
        ["Feature ID", "Feature Name", "Model", "Sales Unit"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "LT1STDLDMC00", "per eNodeB"],
    ])
    s.h3("5.3.2.1  Prerequisite")
    s.table([
        ["Function Name", "Function Switch", "Reference", "Description"],
        ["Downlink D-MIMO", "DmimoJTSwitch of CellAlgoSwitch.DMIMOAlgoSwitch", "D-MIMO (TDD)", "Inter-eNodeB DL D-MIMO depends on downlink D-MIMO."],
    ])
    s.para("Mutually exclusive functions: None. Function impacts: None.")
    s.h2("5.3.3  Hardware")
    s.bullets([
        "3900 and 5900 series base stations.",
        "BBP and co-processing: same as chapter 4 Boards.",
        "Main control board MUST be UMPTe / UMPTg / UMPTe+UMPTg.",
        "The D-MIMO cluster and co-processing board MUST be configured on the BBU where layer 2 of the SFN cell is configured.",
        "RF modules and cells: same as downlink D-MIMO.",
    ])
    s.h2("5.3.4  Networking / cluster IDs")
    s.para("Requires downlink D-MIMO networking (4.3.4). Extra: if multiple cluster IDs for multiple carriers of the same RRU have the same mod 6 value, D-MIMO Calibration Result = Channel calibration failures in inter-eNodeB scenarios.")
    s.h2("5.4.1.1  Data preparation")
    s.para("Deploy by: (1) enable downlink D-MIMO (4.4.1.1); (2) set clock parameters for inter-eNodeB services.")
    s.h3("Table 5-1  Clock parameters")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Clock Working Mode", "TASM.MODE", "Must NOT be FREE for top-level USUs. MAY be FREE for lower-level USUs."],
        ["Locked System Clock Source (1-level USU)", "TASM.SYSCLKSRC", "LOCAL for first-level USUs. INTER_SYSCLK for the BBU."],
        ["Locked System Clock Source (2-level USU)", "TASM.SYSCLKSRC", "LOCAL for second-level USU. INTER_SYSCLK for first-level USUs. INTER_SYSCLK for the BBU."],
    ])
    s.note("Inter-eNodeB D-MIMO is NOT supported when there is a backup for the first-level USU or second-level USU.", "WARNING")
    s.h2("5.4.1.2  MML")
    s.para("Enable downlink D-MIMO (see 4.4.1.2). Set clock parameters per USU3910-based Multi-BBU Interconnection.")
    s.h2("5.4.2  Activation verification")
    s.para("Same as 4.4.2. Extra: if DSP DMIMOCALIBRATION shows route search failed, DSP INREC cannot be executed on the secondary BBU to query inter-RRU calibration interference power. Otherwise DSP INREC can be executed on the secondary BBU.")
    s.h2("5.4.3  Network monitoring")
    s.para("Same as 4.4.3.")
    s.h2("5.4.4  Possible issues")
    s.table([
        ["DSP DMIMOCLUSTERCELL status", "Action"],
        ["Clock exceptions", "DSP SYSCLKSRC: Locked System Clock Source must be Interconnection System Clock AND ESN of the NE where the locked source is located must be the same for all BBUs and USUs. If not, check clock configurations."],
        ["Incorrect configurations or unavailable licenses", "Check whether the inter-eNodeB link is functional."],
        ["Other faults", "See 4.4.4 Possible Issues."],
    ])
    s.footer_end()


def ch6_principles(book):
    s = book.sheet("06-1 Inter-Cell Principles", "6  Inter-Cell D-MIMO  —  Principles and Network Analysis", "Source pages 57–61. Feature TDLEOFD-130501.")
    s.h2("6.1  Principles")
    s.para(
        "Inter-cell D-MIMO enables RRUs with overlapping coverage and corresponding physical cells under the same BBU to form a D-MIMO cluster. There are multiple PCIs within the cluster. Outdoor macro-macro: cluster allows common cells only — not a mix of common and SFN cells. LampSite: compatible with SFN, but each pRRU must serve only one cell."
    )
    s.para("Two types:")
    s.bullets([
        "Intra-BBP inter-cell D-MIMO: Layers 1 and 2 of the cells share the same BBP (Figure 6-1).",
        "Inter-BBP inter-cell D-MIMO: cells are set up on different BBPs in the same BBU (Figure 6-2).",
    ])
    s.figure(f"{FIG}/fig_6_1_intra_bbp_intercell.png", "Figure 6-1  Intra-BBP inter-cell D-MIMO (cells on one UBBP, LCOP in slot)")
    s.figure(f"{FIG}/fig_6_2_inter_bbp_intercell.png", "Figure 6-2  Inter-BBP inter-cell D-MIMO (cells on different UBBPs, same BBU)")
    s.para(
        "To reduce inter-cell interference, PCI mod 3 values need to be aligned for cells in the cluster. If not aligned, CRSs of neighboring cells face data REs of the local cell. Alignment separates CRS and data REs, maximizing JT gains."
    )
    s.figure(f"{FIG}/fig_6_3_pci_mod3_alignment.png", "Figure 6-3  Network reconstruction for PCI mod 3 alignment (staggering vs alignment)")
    s.para("Suitable in outdoor macro-macro and LampSite, especially outdoor hotspots (dense urban, CBD, campus). An inter-cell cluster has resources of multiple logical cells and can serve more UEs than intra-cell D-MIMO.")
    s.h3("Related optimization (eRAN TDD 13.1)")
    s.para("Coordinated scheduling experience optimization: DL_COORD_SCH_EXP_OPT_SW of CellAlgoSwitch.DlSchExtSwitch. Increases coordinated-scheduling proportion and CEU perceived rate. Recommended when CEU perceived rate is low.")
    s.h2("6.2.1  Benefits / most beneficial scenarios")
    s.bullets([
        "Significant overlap, UE speed < 5 km/h, short ISD (≤ 300 m outdoor; LOS between pRRUs in LampSite).",
        "Macro-macro: more than 20% of intra-frequency coverage within the 6 dB scope overlaps, PRB usage > 10%, proportion of BF UEs in macro cells > 20%.",
    ])
    s.para("NOT recommended:")
    s.bullets([
        "High-speed scenarios (D-MIMO depends on beamforming).",
        "Too many UEs (e.g. > 400) — only short-SRS-period UEs can enter BF.",
        "Low coverage overlap rate.",
    ])
    s.para("Benefits: increases average DL user-perceived rate and DL CEU-perceived rate in typical outdoor macro-macro or LampSite. CEU throughput in full-buffer tests increases if the coordinating cell has remaining PRBs for SU-JT, OR no remaining PRBs but meets MU-JT triggering conditions. Otherwise CEU throughput equals normal.")
    s.h2("6.2.2  Impacts")
    s.bullets([
        "Reported 4-bit CQI: decreases in light-load macro-macro because CRS SINR falls after CRS alignment; increases if dual-stream BF proportion grows (TM8 reports one-stream CQI vs TM3 rank2 two-stream).",
        "Slightly decreases RAR, RRC setup, RRC reestablishment, and handover success rates; slightly increases UE access delay. Average DL VoLTE delay and call connection delay increase by less than 10 ms.",
        "Lab full-buffer + HARQ on PUSCH: MU-JT impact on DL cell throughput can exceed common scheduling. Live networks without full-buffer are not affected this way (3GPP TS36.213 §7.3).",
    ])
    s.footer_end()


def ch6_requirements(book):
    s = book.sheet("06-2 Inter-Cell Requirements", "6.3  Inter-Cell Requirements", "Source pages 61–72. Complete these gates before MML.")
    s.h2("6.3.1  Licenses")
    s.table([
        ["Feature ID", "Feature Name", "Model", "Sales Unit"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "LT1SICDLDM00", "per Cell"],
    ])
    s.h3("6.3.2.1  Prerequisite functions")
    s.table([
        ["Function", "Switch", "Reference", "Description"],
        ["Single-stream beamforming", "BfSwitch of CellAlgoSwitch.BfAlgoSwitch", "Beamforming (TDD)", "Outdoor macro-macro: prerequisite. LampSite: license not required but BfSwitch MUST be selected."],
        ["MU beamforming", "MuBfSwitch of CellAlgoSwitch.MuBfAlgoSwitch", "Beamforming (TDD)", "Outdoor macro-macro: prerequisite. LampSite: license not required but MuBfSwitch MUST be selected."],
    ])
    s.h3("6.3.2.2  Mutually exclusive functions")
    s.table([
        ["Function", "Switch", "Note"],
        ["Extended CP", "Cell.UlCyclicPrefix / DlCyclicPrefix", "Inter-cell D-MIMO does not support extended CP."],
        ["High speed mobility", "Cell.HighSpeedFlag", "Does not apply to high-speed scenarios."],
        ["Out-of-band relay", "CellAlgoSwitch.RelaySwitch", "Does not apply."],
        ["Massive MIMO introduction", "SrvBasedSRSAdjAlgo of SRSCfg.SrsCfgPolicySwitch", "Does not apply to massive MIMO scenarios."],
        ["Intra-frequency split (soft split)", "—", "When an intra-frequency soft split cell is added to a cluster, inter-cell DL D-MIMO cannot take effect. Inter-frequency soft split is compatible."],
        ["SFN", "—", "Outdoor: only common cells, not mix with SFN. LampSite: compatible with SFN but each pRRU serves only one logical cell."],
        ["Uplink SU-MIMO", "CellAlgoSwitch.UlSuMimoAlgoSwitch", "Inter-cell D-MIMO cannot take effect when UL SU-MIMO is on."],
        ["Dynamic Power Sharing Between LTE Carriers", "LTE_DYN_POWER_SHARING_SW", "Cannot be enabled together."],
        ["Turbo beamforming", "TurboBfSwitch of MuBfAlgoSwitch", "Cannot be enabled together."],
        ["Enhanced coordinated scheduling based power control", "CellCspcPara.CelleCspcSwitch", "Cannot be enabled together."],
        ["SRS interference avoidance", "WTTxSRSIntrfAvoidanceSw", "Cannot be enabled together."],
        ["Energy saving based on proactive scheduling", "SymbolPwrSaving.TrigBndlSchDlAvgPrbThld", "Cannot be enabled together."],
        ["RF channel dynamic muting", "RF_CHN_DYN_MUTING_SW", "None (listed exclusive table)"],
        ["Dynamic TM3R2↔TM9R4 / dual-BF threshold adapt", "BfMimoAlgoOptSwitch options", "Does not take effect if InterCellDmimoJTSwitch is selected."],
        ["Intelligent Multi-Beam of 8T8R", "—", "None"],
        ["Adaptive MU beamforming", "CellBf.AdaptMubfStartThld ≠ 255", "None"],
        ["MuteUpptsSymbForSrsSw", "SRSCfg.SrsCfgPolicySwitch", "Cannot be on together with inter-cell D-MIMO."],
        ["Single-cell intelligent PC energy saving", "INTEL_INTRA_CELL_PC_SW", "None"],
    ])
    s.h3("6.3.2.3  Function impacts")
    s.table([
        ["Function", "Impact"],
        ["DRX", "No SRS in sleep → BF weights stale → D-MIMO performance down."],
        ["Adaptive switching BF↔MIMO", "Light load: TMA for single-link. Heavy load: use cell-level TMA when multi-layer pairing available."],
        ["DL CoMP", "D-MIMO used for CEU interference suppression; Intra-eNodeB DL CoMP counters measured in a D-MIMO way. Reserved SRS periods 1/2 must be the same for cells in the same CoMP cluster. Intra-eNodeB DL CoMP does not take effect where inter-cell D-MIMO has taken effect (GUI may show both for ≤1 minute)."],
        ["Downlink frequency selective scheduling", "When both on, D-MIMO JT is performed preferentially."],
        ["Uplink coordinated AMC / inter-eNB coordinated AMC", "Reserved SRS period 1/2 must match across cells in the same cluster."],
        ["Downlink CA", "SCells of a UE cannot be selected for D-MIMO JT for that UE."],
        ["RF channel intelligent shutdown / intelligent carrier power-off / low power / cell switch-off auto co-coverage / multi-carrier coordinated ES / SOC low-power", "Cal fails and/or cell cannot be coordinating cell → cluster disabled."],
        ["DL 2-layer MIMO based on TM9", "JT only if ALL cells in the cluster have the same CSI-RS time-frequency position."],
        ["RAN sharing with common carrier", "Cells in the cluster must belong to the same operator. JT only for UEs of that operator; in inter-cell, JT by multiple cells only for primary-operator UEs."],
        ["Adaptation BF↔MIMO with TM4", "JT cannot be performed for UEs in TM3 or TM4."],
        ["Downlink FDD+TDD CA", "Inter-cell downlink D-MIMO does not work in SCells. (Editorial change in eRAN22.1 Draft A.)"],
    ])
    s.h2("6.3.3  Hardware")
    s.bullets([
        "3900/5900 macro; DBS3900/DBS5900 LampSite.",
        "BBP: UBBPg, UBBPd6, UBBPd9, or UBBPe series. UBRlb does not support D-MIMO. FDD+TDD UBBPg is incompatible with inter-cell D-MIMO. Cells can be on different BBPs but same BBU. BBU3910A not supported.",
        "Dedicated cooperative processing board (LCOP) required. Each LCOP: max 288 antennas, 36 cells, 12 D-MIMO clusters.",
        "LCOP uses ENodeBFrameOffset.TddFrameOffset. Each cell's CellFrameOffset.FrameOffset must match the base station frame offset.",
        "RF table same as intra-cell (RRU3252/3256/3278/3279, pRRU3901/3911). Same split/NR/combine/cascade/ring/load-share/20 km/3-carrier/FDD+TDD cascade constraints.",
    ])
    s.h3("Cells")
    s.bullets([
        "Bandwidth 10 or 20 MHz.",
        "Outdoor: bands 38, 40, 42, 43.",
        "LampSite: 2.3 GHz, 2.5 GHz, and band 41.",
        "A cell configured with simulated loads cannot be serving or coordinating cell of a D-MIMO UE.",
        "LampSite: each pRRU for a single cell; same RHUB or two neighboring RHUBs on same CPRI; no extender.",
    ])
    s.h2("6.3.4  Networking / cluster planning")
    s.bullets([
        "Plan with Huawei D-MIMO cluster planning tool.",
        "Each cluster supports at least two cells in macro networking.",
        "Same cluster: same bandwidth, frequency, UL-DL subframe config, special subframe config, and number of CRS ports.",
        "Same PCI mod 3 value. Plan PCIs of cluster cells AND neighbors with Huawei PCI planning tool.",
        "Macro-macro: up to 4 macro cells (max 32 antennas).",
        "LampSite: 3 to 6 pRRUs; multi-carrier: same LTE TDD carrier count and bandwidth for all pRRUs in the cluster.",
    ])
    s.footer_end()


def ch6_config(book):
    s = book.sheet("06-3 Inter-Cell Data Prep", "6.4.1.1  Inter-Cell Data Preparation", "Source pages 73–78.")
    s.h2("Table 6-1  CoProcRes MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["Coordinate Process Resource ID", "CoProcRes.CoProcResId", "0–12"],
        ["BaseBand Equipment ID", "CoProcRes.BaseBandEqmId", "ID of baseband equipment of the dedicated co-processing board. 0–23."],
        ["Work Mode", "CoProcRes.WorkMode", "Select COORDINATING_PROCESSING and DESELECT BASEBAND_PROCESSING if inter-cell D-MIMO is supported."],
        ["Bundling Cluster Type", "CoProcRes.BundlingClusterType", "DMIMO if only inter-cell D-MIMO. ADAPTIVE only if D-MIMO and other coordination features."],
    ])
    s.note(
        "Inter-cell D-MIMO can only support the dedicated co-processing board, which does not support baseband processing — do NOT select BASEBAND_PROCESSING. If COORDINATING_PROCESSING is selected on a BBP that does not support coordinated processing, ALM-26245 Configuration Data Inconsistency is generated; only universal baseband processing of that BBP takes effect so cell services are not affected.",
        "NOTICE",
    )
    s.h2("Table 6-2  DMIMOCluster MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["D-MIMO Cluster ID", "DMIMOCluster.DMIMOClusterId", "0–65535. Even if >4 pRRUs LampSite. Unique mod 6 per RRU multi-carrier. Unique within BBU."],
        ["Coordinate Process Resource ID", "DMIMOCluster.CoProcResId", "255 or 0–12. 255 = none specified."],
        ["Bf Weight Normalized Mode", "DMIMOCluster.BfWeightNormalizeMode", "NEBF commercial (edge). PEBF center/medium."],
    ])
    s.h2("Table 6-3  DMIMOClusterCell MO")
    s.table([
        ["Parameter Name", "Parameter ID", "Setting Notes"],
        ["D-MIMO Cluster ID", "DMIMOClusterCell.DMIMOClusterId", "0–65535"],
        ["Cell ID", "DMIMOClusterCell.CellId", "0–255"],
        ["eNodeB ID", "DMIMOClusterCell.eNodeBId", "0–1048575"],
        ["MCC", "DMIMOClusterCell.Mcc", "000–999"],
        ["MNC", "DMIMOClusterCell.Mnc", "00–99 or 000–999"],
    ])
    s.note(
        "Inter-cell cluster does not allow SFN cells. MPRU aggregation cells are not allowed in outdoor macro-macro. LampSite allows multiple MPRU aggregation cells but a pRRU is configured for a single cell. One cell cannot join different clusters. Add by primary or secondary operator ID, not both (recommend primary).",
        "NOTICE",
    )
    s.h2("Table 6-4  CellAlgoSwitch")
    s.table([
        ["Parameter Name", "Parameter ID", "Option", "Setting Notes"],
        ["BF algorithm switch", "CellAlgoSwitch.BfAlgoSwitch", "BfSwitch", "Select this option."],
        ["MUBF Algorithm Switch", "CellAlgoSwitch.MuBfAlgoSwitch", "MuBfSwitch", "Select this option."],
        ["D-MIMO Algorithm Switch", "CellAlgoSwitch.DMIMOAlgoSwitch", "InterCellDmimoJTSwitch", "Select this option."],
        ["Enhanced Channel Calibration Switch", "CellAlgoSwitch.EnhChnCalSwitch", "QUICK_CHN_CAL_SWITCH (recommended)", "Turn on when D-MIMO is enabled."],
        ["Interference avoidance switch", "CellAlgoSwitch.AvoidInterfSwitch", "AvoidCrsConflictInterfSw", "Select when PCI mod 3 values are aligned."],
        ["DL Scheduling Extension Switch", "CellAlgoSwitch.DlSchExtSwitch", "DL_COORD_SCH_EXP_OPT_SW", "Select when inter-cell D-MIMO is enabled."],
    ])
    s.h2("Tables 6-5 to 6-11  other MOs")
    s.table([
        ["MO / Parameter", "ID", "Setting when inter-cell D-MIMO is ON"],
        ["CellPcAlgo — DM SRS PC SINR offset", "CellPcAlgo.DMSrsPcSinrOffset", "5 (0 when disabled)"],
        ["DMIMOAlgo — D-MIMO A3 RSRP Offset", "DMIMOAlgo.DmimoA3RsrpOffset", "-12 (actual –6 dB)"],
        ["DMIMOAlgo — JT Cell RSRP Diff Thld", "DMIMOAlgo.DmimoJtCellRsrpDiffThld", "-20 (actual –10 dB)"],
        ["SRSCfg — Reserved Resource SRS Period 1", "SRSCfg.RsvResSrsPeriod1", "10ms (same for all cells in cluster/CoMP/AMC)"],
        ["SRSCfg — Reserved Resource SRS Period 2", "SRSCfg.RsvResSrsPeriod2", "20ms"],
        ["CellCounterParaGroup — Edge User A3 Offset", "CellCounterParaGroup.EdgeUserA3Offset", "-12"],
        ["CellCounterParaGroup — Cell Counter Algo Switch", "CellCounterParaGroup.CellCounterAlgoSwitch", "Select BasedA3EdgeUserSwitch"],
        ["NCellSrsMeasPara — SRS Auto NCell Meas", "NCellSrsMeasPara.SrsAutoNCellMeasSwitch", "ON"],
        ["NCellSrsMeasPara — A3 Offset", "NCellSrsMeasPara.NCellSrsMeasA3Offset", "-12"],
        ["NCellSrsMeasPara — NCell Meas Switch", "NCellSrsMeasPara.NCellMeasSwitch", "Select NCellSrsTimeMeasSwitch"],
        ["CellPdcchAlgo — PDCCH Agg Lvl Adapt Strategy", "CellPdcchAlgo.PDCCHAggLvlAdaptStrage", "STRATEGYBASEDONCAPACITY"],
        ["CellCqiAdjAlgo — CQI Adj Schedule Count Thld", "CellCqiAdjAlgo.CellCqiAdjSchCntThld", "50"],
    ])
    s.h2("Optimization — SRS resource allocation (LampSite inter-cell)")
    s.para("SRSCfg.SrsResOptSwitch. All of: inter-cell D-MIMO enabled; UBBP manually bound; TddSrsCfgMode=ACCESS_ENHANCED; Cell.TxRxMode=2T2R. High-risk. Requires UBBPe/UBBPd9. UBBPd4/LBBPd: ALM-29243 / 29240.")
    s.footer_end()


def ch6_mml(book):
    s = book.sheet("06-4 Inter-Cell MML", "6.4.1.2–6.4.1.3  Inter-Cell MML and MAE", "Source pages 78–81. Example: three macro cells LocalCellId 0/1/2, PCI 3/6/9 (mod 3 aligned).")
    s.para("Complete 6.3 Requirements first. Confirm Service Interrupted After Modification / Caution in parameter reference.")
    s.mml("Baseband + cells (8T8R, band 38, SA2/SSP7, same EARFCN, MultiRruCellFlag=FALSE)",
"""ADD BASEBANDEQM: BASEBANDEQMID=0, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=3;
ADD BASEBANDEQM: BASEBANDEQMID=1, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=4;
ADD CELL: LocalCellId=0, CellName="huawei", FreqBand=38, UlEarfcnCfgInd=NOT_CFG, DlEarfcn=37850, UlBandWidth=CELL_BW_N100, DlBandWidth=CELL_BW_N100, CellId=1, PhyCellId=3, FddTddInd=CELL_TDD, SubframeAssignment=SA2, SpecialSubframePatterns=SSP7, EuCellStandbyMode=ACTIVE, RootSequenceIdx=10, MultiRruCellFlag=BOOLEAN_FALSE, TxRxMode=8T8R;
ADD CELL: LocalCellId=1, CellName="huawei", FreqBand=38, UlEarfcnCfgInd=NOT_CFG, DlEarfcn=37850, UlBandWidth=CELL_BW_N100, DlBandWidth=CELL_BW_N100, CellId=2, PhyCellId=6, FddTddInd=CELL_TDD, SubframeAssignment=SA2, SpecialSubframePatterns=SSP7, EuCellStandbyMode=ACTIVE, RootSequenceIdx=10 , MultiRruCellFlag=BOOLEAN_FALSE, TxRxMode=8T8R;
ADD CELL: LocalCellId=2, CellName="huawei", FreqBand=38, UlEarfcnCfgInd=NOT_CFG, DlEarfcn=37850, UlBandWidth=CELL_BW_N100, DlBandWidth=CELL_BW_N100, CellId=3, PhyCellId=9, FddTddInd=CELL_TDD, SubframeAssignment=SA2, SpecialSubframePatterns=SSP7, EuCellStandbyMode=ACTIVE, RootSequenceIdx=10 , MultiRruCellFlag=BOOLEAN_FALSE, TxRxMode=8T8R;
ADD CELLOP: LocalCellId=0, TrackingAreaId=0;
ADD CELLOP: LocalCellId=1, TrackingAreaId=0;
ADD CELLOP: LocalCellId=2, TrackingAreaId=0;
ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=0, BaseBandEqmId=1;
ADD EUCELLSECTOREQM: LocalCellId=1, SectorEqmId=1, BaseBandEqmId=1;
ADD EUCELLSECTOREQM: LocalCellId=2, SectorEqmId=2, BaseBandEqmId=1;
ADD EUCELLPRIBBEQM:LocalCellId=0,PriBaseBandEqmId=1;
ADD EUCELLPRIBBEQM:LocalCellId=1,PriBaseBandEqmId=1;
ADD EUCELLPRIBBEQM:LocalCellId=2,PriBaseBandEqmId=1;""")
    s.mml("Co-processing (dedicated LCOP — COORDINATING_PROCESSING on, BASEBAND_PROCESSING off) + cluster membership",
"""ADD COPROCRES: CoProcResId=0, BaseBandEqmId=0, BundlingClusterType=DMIMO, WorkMode=COORDINATING_PROCESSING-1&BASEBAND_PROCESSING-0;
ADD DMIMOCLUSTER: DMIMOClusterId=0, CoProcResId=0;
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=1, eNodeBId=0, Mcc="460", Mnc="01";
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=2, eNodeBId=0, Mcc="460", Mnc="01";
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=3, eNodeBId=0, Mcc="460", Mnc="01";""")
    s.mml("Per-cell algorithm, SRS neighbor meas, DMIMOAlgo, reserved SRS, CEU counters, neighbors, PDCCH, CQI",
"""MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=5;
MOD CELLPCALGO: LocalCellId=1, DMSrsPcSinrOffset=5;
MOD CELLPCALGO: LocalCellId=2, DMSrsPcSinrOffset=5;
MOD NCELLSRSMEASPARA: LocalCellId=0, SrsAutoNCellMeasSwitch=ON, NCellSrsMeasA3Offset=-12, NCellMeasSwitch=NCellSrsTimeMeasSwitch-1;
MOD NCELLSRSMEASPARA: LocalCellId=1, SrsAutoNCellMeasSwitch=ON, NCellSrsMeasA3Offset=-12, NCellMeasSwitch=NCellSrsTimeMeasSwitch-1;
MOD NCELLSRSMEASPARA: LocalCellId=2, SrsAutoNCellMeasSwitch=ON, NCellSrsMeasA3Offset=-12, NCellMeasSwitch=NCellSrsTimeMeasSwitch-1;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-1, MuBfAlgoSwitch=MuBfSwitch-1, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-1, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-1, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-1;
MOD CELLALGOSWITCH: LocalCellId=1, BfAlgoSwitch=BfSwitch-1, MuBfAlgoSwitch=MuBfSwitch-1, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-1, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-1, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-1;
MOD CELLALGOSWITCH: LocalCellId=2, BfAlgoSwitch=BfSwitch-1, MuBfAlgoSwitch=MuBfSwitch-1, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-1, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-1, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-1;
MOD CELLALGOSWITCH: LocalCellId=0, AvoidInterfSwitch=AvoidCrsConflictInterfSw-1;
MOD CELLALGOSWITCH: LocalCellId=1, AvoidInterfSwitch=AvoidCrsConflictInterfSw-1;
MOD CELLALGOSWITCH: LocalCellId=2, AvoidInterfSwitch=AvoidCrsConflictInterfSw-1;
MOD DMIMOALGO: LocalCellId=0, DmimoA3RsrpOffset=-12, DmimoJtCellRsrpDiffThld=-20;
MOD DMIMOALGO: LocalCellId=1, DmimoA3RsrpOffset=-12, DmimoJtCellRsrpDiffThld=-20;
MOD DMIMOALGO: LocalCellId=2, DmimoA3RsrpOffset=-12, DmimoJtCellRsrpDiffThld=-20;
MOD SRSCFG: LocalCellId=0, RsvResSrsPeriod1=10ms, RsvResSrsPeriod2=20ms;
MOD SRSCFG: LocalCellId=1, RsvResSrsPeriod1=10ms, RsvResSrsPeriod2=20ms;
MOD SRSCFG: LocalCellId=2, RsvResSrsPeriod1=10ms, RsvResSrsPeriod2=20ms;
MOD CELLCOUNTERPARAGROUP: LocalCellId=0,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-1, EdgeUserA3Offset=-12;
MOD CELLCOUNTERPARAGROUP: LocalCellId=1,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-1, EdgeUserA3Offset=-12;
MOD CELLCOUNTERPARAGROUP: LocalCellId=2,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-1, EdgeUserA3Offset=-12;
ADD EUTRANINTRAFREQNCELL:LocalCellId=0,Mcc="460",Mnc="01",eNodeBId=0,CellId=2;
ADD EUTRANINTRAFREQNCELL:LocalCellId=1,Mcc="460",Mnc="01",eNodeBId=0,CellId=1;
ADD EUTRANINTRAFREQNCELL:LocalCellId=0,Mcc="460",Mnc="01",eNodeBId=0,CellId=3;
ADD EUTRANINTRAFREQNCELL:LocalCellId=2,Mcc="460",Mnc="01",eNodeBId=0,CellId=1;
ADD EUTRANINTRAFREQNCELL:LocalCellId=1,Mcc="460",Mnc="01",eNodeBId=0,CellId=3;
ADD EUTRANINTRAFREQNCELL:LocalCellId=2,Mcc="460",Mnc="01",eNodeBId=0,CellId=2;
MOD CELLPDCCHALGO: LocalCellId=0, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCAPACITY;
MOD CELLPDCCHALGO: LocalCellId=1, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCAPACITY;
MOD CELLPDCCHALGO: LocalCellId=2, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCAPACITY;
MOD CELLCQIADJALGO: LocalCellId=0, CellCqiAdjSchCntThld = 50;
MOD CELLCQIADJALGO: LocalCellId=1, CellCqiAdjSchCntThld = 50;
MOD CELLCQIADJALGO: LocalCellId=2, CellCqiAdjSchCntThld = 50;""")
    s.h2("Deactivation command examples")
    s.mml("Deactivation (inter-cell) — restore offsets shown in the source",
"""MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=0;
MOD CELLPCALGO: LocalCellId=1, DMSrsPcSinrOffset=0;
MOD CELLPCALGO: LocalCellId=2, DMSrsPcSinrOffset=0;
MOD NCELLSRSMEASPARA: LocalCellId=0, SrsAutoNCellMeasSwitch=OFF, NCellSrsMeasA3Offset=-20, NCellMeasSwitch=NCellSrsTimeMeasSwitch-0;
MOD NCELLSRSMEASPARA: LocalCellId=1, SrsAutoNCellMeasSwitch=OFF, NCellSrsMeasA3Offset=-20, NCellMeasSwitch=NCellSrsTimeMeasSwitch-0;
MOD NCELLSRSMEASPARA: LocalCellId=2, SrsAutoNCellMeasSwitch=OFF, NCellSrsMeasA3Offset=-20, NCellMeasSwitch=NCellSrsTimeMeasSwitch-0;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-0, MuBfAlgoSwitch=MuBfSwitch-0, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-0, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-0, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-0;
MOD CELLALGOSWITCH: LocalCellId=1, BfAlgoSwitch=BfSwitch-0, MuBfAlgoSwitch=MuBfSwitch-0, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-0, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-0, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-0;
MOD CELLALGOSWITCH: LocalCellId=2, BfAlgoSwitch=BfSwitch-0, MuBfAlgoSwitch=MuBfSwitch-0, DMIMOAlgoSwitch=InterCellDmimoJTSwitch-0, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-0, DlSchExtSwitch=DL_COORD_SCH_EXP_OPT_SW-0;
MOD CELLALGOSWITCH: LocalCellId=0, AvoidInterfSwitch=AvoidCrsConflictInterfSw-0;
MOD CELLALGOSWITCH: LocalCellId=1, AvoidInterfSwitch=AvoidCrsConflictInterfSw-0;
MOD CELLALGOSWITCH: LocalCellId=2, AvoidInterfSwitch=AvoidCrsConflictInterfSw-0;
RMV DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=0, eNodeBId=0, Mcc="100", Mnc="01";
RMV DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=1, eNodeBId=0, Mcc="100", Mnc="01";
RMV DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=2, eNodeBId=0, Mcc="100", Mnc="01";
RMV DMIMOCLUSTER: DMIMOClusterId=0;
MOD SRSCFG: LocalCellId=0, RsvResSrsPeriod2=10ms;
MOD SRSCFG: LocalCellId=1, RsvResSrsPeriod2=10ms;
MOD SRSCFG: LocalCellId=2, RsvResSrsPeriod2=10ms;
MOD CELLCOUNTERPARAGROUP: LocalCellId=0,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-0, EdgeUserA3Offset=-13;
MOD CELLCOUNTERPARAGROUP: LocalCellId=1,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-0, EdgeUserA3Offset=-13;
MOD CELLCOUNTERPARAGROUP: LocalCellId=2,CellCounterAlgoSwitch=BasedA3EdgeUserSwitch-0, EdgeUserA3Offset=-13;
MOD CELLPDCCHALGO: LocalCellId=0, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCOVERAGE;
MOD CELLPDCCHALGO: LocalCellId=1, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCOVERAGE;
MOD CELLPDCCHALGO: LocalCellId=2, PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCOVERAGE;
MOD CELLCQIADJALGO: LocalCellId=0, CellCqiAdjSchCntThld = 0;
MOD CELLCQIADJALGO: LocalCellId=1, CellCqiAdjSchCntThld = 0;
MOD CELLCQIADJALGO: LocalCellId=2, CellCqiAdjSchCntThld = 0;""")
    s.note("Source deactivation RMV DMIMOCLUSTERCELL uses CellId 0/1/2 and MCC 100/01, while activation used CellId 1/2/3 and MCC 460/01. Use the live CellId/MCC/MNC/eNodeBId of the cluster you actually added.", "NOTICE")
    s.para("MAE-Deployment: see Feature Configuration Using the MAE-Deployment. OSS GUIs may differ.")
    s.footer_end()


def ch6_verify(book):
    s = book.sheet("06-5 Inter-Cell Verify", "6.4.2–6.4.4  Inter-Cell Verification, Monitoring, Issues", "Source pages 81–90. MAE GUI is the same as intra-cell (re-embedded).")
    s.h2("6.4.2  Activation Verification")
    s.step(1, "DSP CELL and LST CELL",
           "Cells working properly if Cell instance state = Normal AND Work Status = Normal for all RRUs.")
    s.step(2, "DSP CELLCALIBRATION",
           "Channel calibration successful if Calibration Result = Success.")
    s.step(3, "DSP DMIMOCLUSTERCELL",
           "Cluster working properly if DMIMO Cluster Cell Status = Normal.")
    s.step(4, "DSP DMIMOCALIBRATION",
           "Succeeded → SU-JT and MU-JT available. Failed → only non-coherent joint transmission (not SU-JT/MU-JT). Cluster can still be Normal even when calibration fails.")
    s.h2("6.4.3  Network Monitoring")
    s.h3("Table 6-12  Pairing counters (no 1Layer.PRB in this table)")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.ChMeas.DMIMO.2Layer.PairPRB", "Average PRBs paired at layer 2"],
        ["L.ChMeas.DMIMO.3Layer.PairPRB", "Average PRBs paired at layer 3"],
        ["L.ChMeas.DMIMO.4Layer.PairPRB", "Average PRBs paired at layer 4"],
        ["L.ChMeas.DMIMO.5Layer.PairPRB", "Average PRBs paired at layer 5"],
        ["L.ChMeas.DMIMO.6Layer.PairPRB", "Average PRBs paired at layer 6"],
        ["L.ChMeas.DMIMO.7Layer.PairPRB", "Average PRBs paired at layer 7"],
        ["L.ChMeas.DMIMO.8Layer.PairPRB", "Average PRBs paired at layer 8"],
    ])
    s.h3("Table 6-13  JT UE proportion — takes effect if L.ChMeas.DMIMO.JT.User.Avg > 0")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.ChMeas.DMIMO.JT.User.Avg", "Number of D-MIMO JT UEs in the downlink"],
        ["L.ChMeas.DMIMO.JTUser.RRU.Avg", "Average working RRUs for D-MIMO JT UEs in the downlink"],
    ])
    s.para("CHR PERIOD_CLUSTER_DMIMO_MR fields same as intra-cell (DmimoClusterID, ClusterAbnormallatency, CalibrationAbnormallatency), default 15 min.")
    s.h3("MAE MIMO (Cell) extra items vs intra-cell")
    s.table([
        ["Monitoring Item", "Unit", "Description"],
        ["Number of successful DMIMO Pairing TM7/TM8/TM9 RB(Num) with N Layers", "Number", "N={2..8}, averaged over TTIs"],
        ["JT User Num", "Number", "Average number of JT UEs in a cell in each TTI"],
        ["JT User Work RRU Num", "Number", "Average working RRUs for JT UEs in the local cell"],
        ["SUJT Schedule RB Num", "Number", "Accumulative PRBs scheduled for JT UEs in SU-JT mode"],
    ])
    s.h3("MAE MIMO (User) — Test Items = TDD DMIMO")
    s.table([
        ["Monitoring Item", "Unit", "Description"],
        ["Number of successful DMIMO pairing RB (Num) with N Layers", "Number", "Range 0–1000"],
        ["DMIMO User Indication", "None", "0 non-JT; 1 JT; 255 invalid"],
        ["JT User Work RRU Num", "Number", "Average working RRUs for JT UEs"],
        ["SUJT Schedule RB Num", "Number", "RBs scheduled for JT UEs in SU-JT"],
    ])
    s.h3("CEU counters (inter-cell names differ — no .JointTransmit suffix)")
    s.table([
        ["Counter Name", "Counter Description"],
        ["L.Thrp.bits.UL.BorderUE", "UL PDCP data volume for CEUs"],
        ["L.Thrp.bits.UL.SmallPkt.BorderUE", "UL PDCP for CEU small packets"],
        ["L.Thrp.Time.UL.RmvSmallPkt.BorderUE", "UL duration for CEU non-small packets"],
        ["L.Thrp.bits.DL.BorderUE", "DL PDCP data volume for CEUs"],
        ["L.Thrp.bits.DL.LastTTI.BorderUE", "DL PDCP last TTI before buffer empty"],
        ["L.Thrp.Time.DL.RmvLastTTI.BorderUE", "DL duration except last TTI"],
    ])
    s.h2("6.4.4  Possible Issues")
    s.para(
        "A D-MIMO cell automatically reverts to a normal cell when a fault related to activation occurs (including RRU antenna rollback). It does NOT revert on a channel calibration fault — then only non-coherent JT is possible."
    )
    s.h3("MAE threshold alarm")
    s.para("Same 5 GUI steps and same counter L.CellSectorEqpt.UNA.Dur.Cali / formula (Period 15 → Threshold 9, Offset 3). Exact screens:")
    s.figure(f"{FIG}/mae_step1_threshold_settings.png", "MAE Threshold Settings — CellSectorEQUIP Performance (same as intra-cell source p.86–87)")
    s.figure(f"{FIG}/mae_step3_add_threshold_object.png", "Add Threshold — Object tab")
    s.figure(f"{FIG}/mae_step4_basic_tab.png", "Add Threshold — Basic tab (CAL Fail, 15 minutes)")
    s.figure(f"{FIG}/mae_step5_advanced_tab.png", "Add Threshold — Advanced tab (L.CellSectorEqpt.UNA.Dur.Cali, 9 / 3)")
    s.h3("Calibration troubleshooting")
    s.para("Same three non-Succeeded values as intra-cell: Exception occurred in internal calibration / Route search failed / Reciprocity calibration failed — same actions.")
    s.h3("Activation status ≠ Normal — extra inter-cell checks")
    s.table([
        ["Status", "Check"],
        ["incorrect configurations or unavailable licenses", "1 switch on 2 RRU count 3 cluster 4 license 5 consistent CRS ports, TDD SRS mode, reserved SRS 1/2, subframe config"],
        ["Limited hardware capacity", "Cells on the same UBBPd9 or UBBPe?"],
        ["Route application failures", "Routing bandwidth"],
        ["Cell abnormal", "DSP CELL"],
        ["Insufficient co-processing resources", "LCOP vs 6.3.3"],
        ["Channel calibration failures", "DSP DMIMOCALIBRATION"],
        ["Cluster ID conflicts", "6.4.1.1 restraints"],
        ["Clusters being established", "Wait"],
        ["Co-processing resources abnormal", "Co-processing on the LCOP?"],
        ["The number of LampSite cells in a cluster must be greater than the specified minimum number to enable D-MIMO", "Two or more cells in the cluster?"],
        ["LampSite cells cannot form a D-MIMO cluster with cells of other types of eNodeBs that have different PCIs", "Only LampSite cells in the cluster?"],
        ["TM8 dual-stream beamforming is configured for micro eNodeBs", "If CellBf.MaxBfRankPara = DUAL_LAYER_BF on micro eNodeB cells, the cluster is abnormal"],
    ])
    s.h3("Table 6-14  eNodeB alarms related to BBPs")
    s.table([
        ["Alarm ID", "Alarm Name", "Action"],
        ["26245", "Configuration Data Inconsistency", "Handle per alarm reference. Check whether the BBP that provides co-processing resources supports centralized control for inter-cell D-MIMO. If not, replace the BBP (see 6.3.3)."],
        ["26203", "Board Software Program Error", "Same as above."],
    ])
    s.para("Others: DmimoJTSwitch selected → possible consecutive block errors in dual-stream BF / disconnects. See Beamforming (TDD).")
    s.footer_end()


def ch7_10(book):
    s = book.sheet("07-10 Params Counters Refs", "7–10  Parameters, Counters, Glossary, Reference Documents", "Source pages 91–95.")
    s.h2("7  Parameters")
    s.para("The following hyperlinked EXCEL files of parameter documents match the software version with which this document is released. Find them from the Man-Machine Interface Reference node in the 3900 & 5900 Series Base Station Product Documentation delivered with that version.")
    s.bullets([
        "Node Parameter Reference — device and transport parameters",
        "eNodeBFunction Parameter Reference — radio access functions (air interface, access, mobility, RRM)",
        "eNodeBFunction Used Reserved Parameter List — reserved parameters in use and disused",
        "Node Used Reserved Parameter List — reserved parameters in use and disused",
    ])
    s.h3("FAQ 1 — parameters related to a feature")
    s.bullets([
        "Open the EXCEL file of parameter reference.",
        "On the Parameter List sheet, filter the Feature ID column. Text Filters → Contains. Enter the feature ID, for example LOFD-001016 or TDLOFD-001016. For this document use TDLEOFD-111505, TDLEOFD-121501, or TDLEOFD-130501.",
        "Click OK. All parameters related to the feature are displayed.",
    ], numbered=True)
    s.h3("FAQ 2 — reserved parameter")
    s.bullets([
        "Open the used reserved parameter list EXCEL.",
        "On Used Reserved Parameter List, use MO, Parameter ID, and BIT columns to locate the reserved parameter (may be only a bit).",
        "View meaning, values, impacts, and product version in which it is activated.",
    ], numbered=True)
    s.h2("8  Counters")
    s.bullets([
        "Node Performance Counter Summary — device and transport counters",
        "eNodeBFunction Performance Counter Summary — radio access counters",
        "eNodeBFunction Used Reserved Counter List — reserved counters in use and disused",
    ])
    s.h3("FAQ — counters related to a feature")
    s.bullets([
        "Open the performance counter reference EXCEL.",
        "On Counter Summary(En), filter Feature ID → Contains → TDLEOFD-111505 / 121501 / 130501.",
        "Click OK.",
    ], numbered=True)
    s.h2("9  Glossary")
    s.para("For the acronyms, abbreviations, terms, and definitions, see Glossary (companion Huawei document). Key terms used in this pack:")
    s.table([
        ["Term", "Meaning in this feature"],
        ["D-MIMO", "Distributed MIMO — centralized scheduling + distributed RRUs"],
        ["CEU", "Cell Edge User"],
        ["JT / SU-JT / MU-JT", "Joint Transmission / Single-User JT / Multi-User JT"],
        ["SFN", "Single Frequency Network (same PCI logical cell)"],
        ["PCI mod 3 alignment", "Cells in an inter-cell cluster share PCI modulo 3 so CRS/data REs align"],
        ["LCOP", "LTE Cooperative Processing board (dedicated co-processing for inter-cell)"],
        ["CoProcRes", "Coordinate Process Resource managed object"],
        ["LampSite / pRRU / RHUB", "Indoor distributed RRU solution"],
        ["TMA", "Transmission Mode Adaptation"],
        ["SRS", "Sounding Reference Signal — used to pick JT UEs and BF weights"],
        ["Cloud BB / USU", "Centralized BBU pool and interconnection switch used by inter-eNodeB D-MIMO"],
    ])
    s.h2("10  Reference Documents")
    s.table([
        ["#", "Document"],
        ["1", "SFN"],
        ["2", "Energy Conservation and Emission Reduction"],
        ["3", "Beamforming (TDD)"],
        ["4", "DRX and Signaling Control"],
        ["5", "Extended CP"],
        ["6", "High Speed Mobility"],
        ["7", "Relay"],
        ["8", "Massive MIMO Basics (TDD)"],
        ["9", "Massive MIMO Enhancements (TDD)"],
        ["10", "TMA (TDD)"],
        ["11", "WTTx Turbo Beamforming (TDD)"],
        ["12", "DL CoMP (TDD)"],
        ["13", "Downlink Scheduling"],
        ["14", "Uplink Coordinated Scheduling"],
        ["15", "RAN Sharing"],
        ["16", "Soft Split Resource Duplex (TDD)"],
        ["17", "Dynamic Power Sharing Between LTE Carriers"],
        ["18", "MIMO"],
        ["19", "LampSite BBU Technical Specifications"],
        ["20", "Cloud BB Overview"],
        ["21", "USU3910-based Multi-BBU Interconnection"],
        ["22", "Massive MIMO Optimization in WTTx Scenarios (TDD)"],
        ["23", "Intelligent Multi-Beam of 8T8R (TDD)"],
        ["24", "Carrier Aggregation"],
        ["25", "Massive MIMO Optimization in WTTx Scenarios (TDD)"],
        ["26", "Green Site"],
    ])
    s.footer_end()
