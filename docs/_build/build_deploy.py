#!/usr/bin/env python3
"""Step-by-step D-MIMO deployment playbook (document-style Excel)."""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from xlsx_style import DocBook

FIG = "/workspace/docs/figures"
OUT = "/workspace/docs/D-MIMO_TDD_eRAN22.1_Step_by_Step_Deployment_Guide.xlsx"


def build():
    book = DocBook(
        title="D-MIMO (TDD) Step-by-Step Deployment Guide",
        subtitle="Field playbook derived in original document order from eRAN22.1 01 Feature Parameter Description",
        doc_code="DEPLOY  |  TDLEOFD-111505 / 121501 / 130501",
    )
    cover(book)
    how_to(book)
    phase0(book)
    phase1(book)
    phase2(book)
    phase3_intra(book)
    phase3_mml(book)
    phase4_interenb(book)
    phase5_intercell(book)
    phase5_mml(book)
    phase6_verify(book)
    phase7_monitor(book)
    phase8_trouble(book)
    phase9_rollback(book)
    checklist(book)
    book.save(OUT)
    print("Wrote", OUT)


def cover(book):
    s = book.sheet("Cover", "D-MIMO (TDD) Step-by-Step Deployment Guide", "Follow sheets left-to-right. Do not skip a gate.")
    s.para("HUAWEI eRAN22.1 01  |  Issue 01 (2026-03-10)", bold=True, color="C7000B")
    s.para("This workbook is the deployment companion to D-MIMO_TDD_eRAN22.1_Feature_Review_and_Summary.xlsx. Every step is taken from the source Feature Parameter Description, in the same order the vendor specified: understand principles → pick networking mode → satisfy requirements → prepare data → run MML/MAE → observe → monitor → handle faults → roll back.")
    s.h2("Which feature are you deploying?")
    s.table([
        ["If your design is…", "Deploy this", "License model", "Start at sheet"],
        ["Several RRUs on ONE BBU already in SFN / pRRU aggregation, SAME PCI", "Intra-cell DL D-MIMO  TDLEOFD-111505  (LT1SDLMIMO00 per cell)", "Commercial", "P1 Mode Select then P3 Intra-Cell"],
        ["RRUs on DIFFERENT BBUs of the SAME Cloud BB, still one logical SFN cell", "Inter-eNodeB DL D-MIMO  TDLEOFD-121501  (LT1STDLDMC00 per eNodeB)", "TRIAL — MoU first, then do chapter 4 PLUS clock", "P1 then P4 Inter-eNodeB"],
        ["Several cells on ONE BBU, DIFFERENT PCIs, PCI mod 3 aligned", "Inter-Cell DL D-MIMO  TDLEOFD-130501  (LT1SICDLDM00 per cell)", "Commercial", "P1 then P5 Inter-Cell"],
    ])
    s.note("Do not mix intra-cell and inter-cell switches on the same cluster. Intra-cell uses DmimoJTSwitch. Inter-cell uses InterCellDmimoJTSwitch.", "WARNING")
    s.note("Source document: activation guidance only. Gains depend on the live scenario. Cluster IDs and (for inter-cell) PCI plans MUST come from Huawei planning tools / professional service.", "NOTICE")
    s.footer_end()


def how_to(book):
    s = book.sheet("00 How to Use", "How to use this playbook", "Print landscape A4. Freeze panes keep the red/navy banner visible.")
    s.h2("Reading order (mandatory)")
    s.table([
        ["Phase", "Sheet", "Exit criterion before you continue"],
        ["P0", "P0 Readiness Gates", "RAT=TDD, software=eRAN22.1, trial MoU if 121501, change-history impacts understood"],
        ["P1", "P1 Mode Select", "Exactly one mode chosen; scenario matches ISD / overlap / speed / LOS"],
        ["P2", "P2 License HW Net", "License loaded, exclusive functions OFF, HW/RF/cluster-ID rules pass"],
        ["P3", "P3 Intra-Cell Steps + P3 Intra-Cell MML", "Only if mode is intra-cell (or as base of inter-eNodeB)"],
        ["P4", "P4 Inter-eNodeB Extra", "Only if trial inter-eNodeB; clock locked; no USU backup"],
        ["P5", "P5 Inter-Cell Steps + P5 Inter-Cell MML", "Only if inter-cell; PCI mod 3 aligned; LCOP dedicated"],
        ["P6", "P6 Activation Verify", "DSP statuses Normal / Succeeded (or known degraded mode)"],
        ["P7", "P7 Network Monitoring", "Pairing counters > 0; MAE alarm configured"],
        ["P8", "P8 Troubleshooting", "Use only if P6/P7 fail"],
        ["P9", "P9 Deactivation Rollback", "Use only to undo"],
        ["CL", "CL Master Checklist", "Sign-off row per site/cluster"],
    ])
    s.h2("Conventions")
    s.bullets([
        "STEP n is a field action. Do it in numeric order inside a phase.",
        "MML blocks are Consolas on dark rows — copy to MAE/LMT. Replace example LocalCellId / MCC / eNodeBId / PCI with live values.",
        "NOTE / NOTICE / WARNING / TRIAL boxes are taken from the source document language.",
        "Figures are the exact charts extracted from D-MIMO (TDD)(eRAN22.1_01).pdf.",
        "Optional steps are marked (Optional) as in the source MML comments.",
    ])
    s.footer_end()


def phase0(book):
    s = book.sheet("P0 Readiness Gates", "Phase 0  —  Readiness gates (before any parameter change)", "Maps to chapters 1–2 of the source.")
    s.step(0.1, "Confirm software and RAT",
           "This pack matches eRAN22.1 01 (2026-03-10) and TDD only. Parameters/alarms/counters/MOs apply only to this release. For another release, use that release’s Feature Parameter Description.",
           "LST SYS / check eRAN version. Abort if not TDD cells.")
    s.step(0.2, "Read the two editorial impacts introduced vs eRAN21.1",
           "Draft A (2025-12-31): (1) inter-cell DL D-MIMO does not work in SCells of downlink FDD+TDD CA. (2) Precise AMC does not take effect for D-MIMO UEs in JT state. Issue 01 itself has no further changes.",
           "If FDD+TDD CA SCells are in the candidate cluster, do not expect inter-cell D-MIMO there.")
    s.step(0.3, "Trial-feature gate (only for TDLEOFD-121501)",
           "Inter-eNodeB DL D-MIMO is a trial feature: not for sale; free trial; MoU required; Huawei not liable; may be auto-invalidated on upgrade if not licensed after conversion.",
           "Stop and obtain MoU + written risk acceptance before any inter-eNodeB cluster is built.")
    s.step(0.4, "License is on the PARENT feature",
           "Subfeatures 11150501/02/04/05 and 13050101/02/03/04 are licensed via TDLEOFD-111505 or TDLEOFD-130501. Sales units: LT1SDLMIMO00 per cell; LT1STDLDMC00 per eNodeB (trial); LT1SICDLDM00 per cell.",
           "DSP LICENSE / LST LICENSE. Confirm cell-activation-affected column in License Control Item Lists.")
    s.step(0.5, "Professional-service gate",
           "Source: this document only guides activation. Gains depend on scenario. Cluster RRU list and cluster ID must be planned — contact Huawei engineers for the D-MIMO cluster planning tool (and PCI planning tool for inter-cell).",
           "Do not proceed with invented cluster IDs.")
    s.footer_end()


def phase1(book):
    s = book.sheet("P1 Mode Select", "Phase 1  —  Choose networking mode (chapter 3 then 4.1 / 5.1 / 6.1)", "Use the exact architecture figures from the source.")
    s.h2("What D-MIMO does (chapter 3)")
    s.para("Site densification creates overlap and cell-edge pain. D-MIMO uses centralized scheduling + distributed RRUs so multiple antennas transmit together. A cluster is a planned set of RRU-generated physical cells. After configuration and switch-on, the eNodeB periodically runs inter-RRU calibration then coherent JT.")
    s.figure(f"{FIG}/fig_3_1a_without_with_dmimo.png", "Figure 3-1 (part)  Without D-MIMO vs With D-MIMO")
    s.figure(f"{FIG}/fig_3_1b_dmimo_architecture.png", "Figure 3-1  Distributed antenna architecture (×10 m to ×100 m)")
    s.figure(f"{FIG}/fig_3_2_dmimo_process.png", "Figure 3-2  Working process you will execute: configure → calibrate → coherent JT")
    s.h2("STEP 1.1  Match scenario to a mode")
    s.table([
        ["Question", "Intra-cell 111505", "Inter-eNodeB 121501 (trial)", "Inter-cell 130501"],
        ["PCI in the cluster?", "SAME PCI (SFN / pRRU aggregation logical cell)", "Same as intra-cell, across BBUs", "DIFFERENT PCIs, same PCI mod 3"],
        ["BBU?", "Same BBU", "Different BBUs, same Cloud BB, full frequency sync", "Same BBU (intra-BBP or inter-BBP)"],
        ["Introduced", "eRAN TDD 11.1", "eRAN TDD 12.1", "eRAN TDD 13.1"],
        ["Switch", "DmimoJTSwitch", "Downlink D-MIMO + Cloud BB clock", "InterCellDmimoJTSwitch"],
        ["Outdoor ISD / overlap", "ISD ≤ 150 m, significant overlap, LOS, UE < 5 km/h", "Same benefits as intra-cell", "ISD ≤ 300 m; >20% overlap in 6 dB; PRB>10%; BF UEs>20%; UE<5 km/h"],
        ["LampSite ISD", "≤ 30 m, severe neighbor interference, LOS, UE<5 km/h", "—", "LOS between pRRUs"],
        ["Do not use if", "High speed, extended CP, Massive MIMO intro, WTTx turbo BF, >exclusive list", "RGPS-host BBU in the inter-eNB cluster; USU backup present", "High speed; >~400 UEs; low overlap; intra-freq soft split cell in cluster"],
        ["RRUs per cluster", "Macro 2–4; LampSite 3–6 pRRUs", "Same as intra-cell plus clock rules", "Macro up to 4 cells (≤32 antennas); LampSite 3–6 pRRUs"],
    ])
    s.figure(f"{FIG}/fig_4_1_intracell_networking.png", "Figure 4-1  Intra-cell / Cloud BB context (source p.21)")
    s.figure(f"{FIG}/fig_5_1_inter_enodeb_networking.png", "Figure 5-1  Inter-eNodeB — RGPS-host BBU is intra-eNodeB-only when RGPS is the network clock")
    s.figure(f"{FIG}/fig_6_1_intra_bbp_intercell.png", "Figure 6-1  Intra-BBP inter-cell")
    s.figure(f"{FIG}/fig_6_2_inter_bbp_intercell.png", "Figure 6-2  Inter-BBP inter-cell (same BBU)")
    s.figure(f"{FIG}/fig_6_3_pci_mod3_alignment.png", "Figure 6-3  Reconstruct PCI mod 3 to alignment BEFORE enabling inter-cell D-MIMO")
    s.step(1.2, "Record the chosen mode on the Master Checklist",
           "Write Feature ID, parent license, PCI rule, BBU rule, cluster ID (from Huawei tool), and site list.",
           "If inter-cell: also record PCI of every member and confirm PCI mod 3 equal. If not equal, stop and replan PCI.")
    s.footer_end()


def phase2(book):
    s = book.sheet("P2 License HW Net", "Phase 2  —  License, exclusive functions, hardware, networking", "Maps to 4.3 / 5.3 / 6.3. Fail any row → do not run activation MML.")
    s.h2("STEP 2.1  Load and confirm license")
    s.table([
        ["Feature ID", "Name", "Model", "Unit", "DSP/LST check"],
        ["TDLEOFD-111505", "DL D-MIMO", "LT1SDLMIMO00", "per Cell", "Enough cell licenses for every SFN/aggregation cell in a cluster"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "LT1STDLDMC00", "per eNodeB", "Per eNodeB; MoU on file"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "LT1SICDLDM00", "per Cell", "Enough cell licenses for every member cell"],
    ])
    s.h2("STEP 2.2  Prerequisite functions ON")
    s.table([
        ["Mode", "Must already be ON"],
        ["Intra-cell", "SFN (SfnUlSchSwitch + SfnDlSchSwitch). Outdoor: single-stream BF license + MU-BF license. LampSite: BfSwitch and MuBfSwitch selected even if those licenses are not required."],
        ["Inter-eNodeB", "Everything for intra-cell PLUS DmimoJTSwitch (downlink D-MIMO) already working on the logical cell."],
        ["Inter-cell", "Outdoor: single-stream BF + MU-BF licenses. LampSite: BfSwitch and MuBfSwitch selected. SFN is NOT a prerequisite (outdoor forbids mixing SFN+common in one cluster)."],
    ])
    s.h2("STEP 2.3  Mutually exclusive functions OFF (union of 4.3.2.2 and 6.3.2.2)")
    s.table([
        ["Turn OFF / do not use", "Applies to", "Parameter / switch"],
        ["Extended CP", "Both", "Cell.UlCyclicPrefix / DlCyclicPrefix"],
        ["High speed mobility", "Both", "Cell.HighSpeedFlag"],
        ["Out-of-band relay", "Both", "CellAlgoSwitch.RelaySwitch"],
        ["Massive MIMO introduction", "Both", "SrvBasedSRSAdjAlgo of SRSCfg.SrsCfgPolicySwitch"],
        ["Uplink SU-MIMO", "Both", "CellAlgoSwitch.UlSuMimoAlgoSwitch"],
        ["WTTx turbo BF / AntSelEnhance BF", "Intra-cell exclusive; turbo also inter-cell exclusive", "TurboBfSwitch / AntSelEnhanceBfSwitch"],
        ["SRS interference avoidance (WTTx)", "Both", "WTTxSRSIntrfAvoidanceSw"],
        ["Energy saving based on proactive scheduling", "Both — cannot combine", "SymbolPwrSaving.TrigBndlSchDlAvgPrbThld"],
        ["RF channel dynamic muting", "Both listed", "RF_CHN_DYN_MUTING_SW"],
        ["MuteUpptsSymbForSrsSw", "Both — cannot combine", "SRSCfg.SrsCfgPolicySwitch option"],
        ["Single-cell intelligent PC ES", "Both listed", "INTEL_INTRA_CELL_PC_SW"],
        ["Adaptive MU-BF (AdaptMubfStartThld ≠ 255)", "Both listed", "CellBf.AdaptMubfStartThld"],
        ["Intelligent Multi-Beam 8T8R", "Both listed", "—"],
        ["Intra-frequency soft split cell in cluster", "Inter-cell", "Soft Split Resource Duplex"],
        ["Dynamic Power Sharing Between LTE Carriers", "Inter-cell cannot combine; intra-cell reduces DPS gain", "LTE_DYN_POWER_SHARING_SW"],
        ["Enhanced coordinated scheduling based PC", "Inter-cell", "CellCspcPara.CelleCspcSwitch"],
    ])
    s.h2("STEP 2.4  Energy-saving / sleep functions that BREAK calibration — leave OFF on cluster cells")
    s.table([
        ["Function", "What happens if left ON"],
        ["RF channel intelligent shutdown", "Inter-RRU calibration fails (intra-cell). Inter-cell: cluster disabled."],
        ["Intelligent power-off of co-coverage carriers / cell switch-off auto co-coverage / multi-carrier coordinated ES", "Cal fails; inter-cell cluster disabled"],
        ["Low power consumption mode / SOC level-based low power", "Cannot be cooperating cell; cal may fail; cluster may disable"],
        ["pRRU deep dormancy (intelligent)", "Downlink D-MIMO no longer takes effect for those pRRUs"],
        ["DRX", "No SRS during sleep → stale BF weights → D-MIMO performance drop (not always exclusive, but treat as risk)"],
    ])
    s.h2("STEP 2.5  Hardware walk (fail closed)")
    s.table([
        ["Check", "Intra-cell", "Inter-cell extra", "Pass?"],
        ["Site type", "3900/5900 macro or DBS3900/5900 LampSite", "Same", "☐"],
        ["Forbidden chassis/boards", "No BBU3910A, no UBRlb, no FDD+TDD BBP as co-processing", "No FDD+TDD UBBPg; LCOP dedicated (no baseband processing on that resource)", "☐"],
        ["BBP family", "UBBPe4/d6/d9/e6 to set up cells", "UBBPg / UBBPd6 / UBBPd9 / UBBPe", "☐"],
        ["Co-processing", "UBBPd6/e6/d9; BBU3900 slot 2 or 3; not dedicated scheduler; antenna/cluster caps 72/6 (UBBPd) or 144/12 (UBBPe)", "LCOP: 288 antennas, 36 cells, 12 clusters; WorkMode COORDINATING_PROCESSING-1 & BASEBAND_PROCESSING-0", "☐"],
        ["Main control (inter-eNB only)", "—", "UMPTe / UMPTg / combination; cluster+co-proc on the BBU that hosts SFN L2", "☐"],
        ["RF", "RRU3252/3256/3278/3279 or pRRU3901/3911 only; no NR RF; no RRU3279 split-4T; no combine/cascade/ring/load-share; ≤20 km; ≤3 carriers", "Same", "☐"],
        ["Cell", "10 or 20 MHz; outdoor 38/40/42/43; LampSite 2.3/2.5/2.6 (intra) or 2.3/2.5/band 41 (inter-cell)", "No simulated-load cells as serving/coordinating", "☐"],
        ["Frame offset", "Cluster offset = co-processing board CellFrameOffset.FrameOffset", "Each cell FrameOffset = ENodeBFrameOffset.TddFrameOffset of LCOP", "☐"],
        ["LampSite cabling", "Same RHUB or two neighboring RHUBs on same CPRI; no extender; no pRRU combination (intra-cell)", "Each pRRU serves one cell", "☐"],
    ])
    s.h2("STEP 2.6  Cluster ID rules (Huawei planning tool output)")
    s.bullets([
        "Unique within a BBU.",
        "Neighboring clusters: IDs must not share the same mod 24.",
        "Multi-carrier: each cluster ID unique mod 6 per RRU. Inter-eNodeB: same-mod-6 on same RRU’s carriers → Calibration Result = Channel calibration failures.",
        "LampSite >4 pRRUs: cluster ID even.",
        "Intra-cell: only ONE cell per cluster (the SFN / aggregation logical cell).",
        "Inter-cell macro: ≥2 cells, ≤4 macro cells; all members same bandwidth, frequency, UL-DL subframe, special subframe, CRS port count, PCI mod 3.",
        "Add cell by primary operator ID (not both primary and secondary).",
    ])
    s.footer_end()


def phase3_intra(book):
    s = book.sheet("P3 Intra-Cell Steps", "Phase 3  —  Intra-cell D-MIMO sequential configuration (4.4.1)", "Do this for TDLEOFD-111505. Also do this as the base of trial inter-eNodeB, then continue to P4.")
    s.h2("Architecture reminder — calibration then JT")
    s.figure(f"{FIG}/fig_3_3_inter_rru_calibration.png", "Figure 3-3  Inter-RRU calibration (why QUICK_CHN_CAL_SWITCH is recommended)")
    s.figure(f"{FIG}/fig_3_4_coherent_jt.png", "Figure 3-4  Coherent JT after calibration")
    s.figure(f"{FIG}/fig_3_5_jt_process.png", "Figure 3-5  Per-TTI SU-JT then MU-JT")
    s.figure(f"{FIG}/fig_3_6_su_jt.png", "Figure 3-6  SU-JT")
    s.figure(f"{FIG}/fig_3_7_mu_jt.png", "Figure 3-7  MU-JT")
    s.figure(f"{FIG}/fig_3_8_mu_jt_example.png", "Figure 3-8  MU-JT resource example — set BfWeightNormalizeMode=NEBF commercially")
    s.step(3.1, "Build / confirm the SFN or pRRU-aggregation cell first",
           "Intra-cell D-MIMO sits on SFN. If the cell is not already SFN, complete SFN feature deployment first (see SFN document). Non-SFN networking must be planned with the D-MIMO cluster planning tool.",
           "LST CELL: Mode of Multi-RRU Cell = SFN or MPRU_AGGREGATION. Cell instance state Normal.")
    s.step(3.2, "Create baseband equipment including the co-processing slot",
           "Example: UBBP slots 0–3, slot 3 = co-processing. BBU3900: co-processing in slot 2 or 3 only.",
           "ADD BASEBANDEQM for each SN. See P3 Intra-Cell MML.")
    s.step(3.3, "Bind sector equipment (macro) or sector equipment groups (LampSite)",
           "Macro example: 3 sector equipment on the SFN cell bound to BB eqm 0–2. LampSite example: 4 groups.",
           "ADD EUCELLSECTOREQM (macro) or ADD EUSECTOREQMGROUP + EUSECTOREQMID2GROUP (LampSite).")
    s.step(3.4, "ADD COPROCRES",
           "CoProcResId 0–12. BaseBandEqmId = co-processing board. BundlingClusterType=DMIMO (or ADAPTIVE if mixed coordination). WorkMode COORDINATING_PROCESSING only if mixed coordination; not required if only D-MIMO.")
    s.step(3.5, "ADD DMIMOCLUSTER",
           "DMIMOClusterId from planning tool (even if >4 pRRUs). CoProcResId matching step 3.4 (255 = none specified). BfWeightNormalizeMode=NEBF for commercial edge-optimized MU-JT.")
    s.step(3.6, "ADD DMIMOCLUSTERCELL (exactly one logical cell)",
           "Map the SFN/aggregation CellId + eNodeBId + MCC + MNC. Primary operator ID recommended.")
    s.step(3.7, "(Optional) Target SRS SINR optimization for BF UEs",
           "MOD CELLPCALGO DMSrsPcSinrOffset=5. May slightly increase RRC Reconfiguration and drop rate.")
    s.step(3.8, "Turn on BF, MU-BF, D-MIMO JT, quick calibration",
           "MOD CELLALGOSWITCH BfSwitch-1, MuBfSwitch-1, DmimoJTSwitch-1, QUICK_CHN_CAL_SWITCH-1. This is the actual activation of intra-cell D-MIMO.")
    s.step(3.9, "(Optional) Advanced Multi-User Coordination",
           "MULTI_UE_COORDINATION_OPT_SW-1 on CoordinationAlgoSwitch. Increases MU-JT pairing and DL capacity in heavy load. See Beamforming (TDD).")
    s.step(3.10, "(Optional) PDCCH BF gain offset = 0",
           "Only when PDCCHAggLvlAdaptStrage = STRATEGYBASEDONCOVERAGE. Enables CCE selection optimization for joint-scheduling BF UEs. Default was -127.")
    s.step(3.11, "(Optional) WaitPairingLayerThd = 30",
           "Only when cell load is heavy AND multi-layer pairing rate is high. Else leave 0. Triggers cell-level TMA so more UEs enter BF for MU pairing if DL PRB usage > internal threshold AND BF UEs/TTI ≥ threshold×0.1.")
    s.step(3.12, "(Optional, LampSite only) SRS resource optimization",
           "HIGH RISK. SrsResOptSwitch only if D-MIMO on, UBBP manually bound, TddSrsCfgMode=ACCESS_ENHANCED, MultiRruCellMode=MPRU_AGGREGATION, board is UBBPe/d6/d9 — never UBBPd4/LBBPd.")
    s.step(3.13, "Other recommended intra-cell optimizations from data-prep tables",
           "TmAccelerationSwitch=INITIAL_ACCESS_TO_BF. Micro-micro only: DlHighLoadSdmaThdOffset=10 if D-MIMO+quick BF, else 5. Not used in macro-macro/macro-micro.")
    s.note("Before any MOD, read Service Interrupted After Modification and Caution in the parameter reference for that MO.", "NOTICE")
    s.footer_end()


def phase3_mml(book):
    s = book.sheet("P3 Intra-Cell MML", "Phase 3  —  Intra-cell MML script (copy/adapt)", "Source 4.4.1.2. Example LocalCellId=0, MCC=100 MNC=01 eNodeBId=0. Replace with live values.")
    s.mml("1) Baseband equipment",
"""ADD BASEBANDEQM: BASEBANDEQMID=0, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=0;
ADD BASEBANDEQM: BASEBANDEQMID=1, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=1;
ADD BASEBANDEQM: BASEBANDEQMID=2, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=2;
ADD BASEBANDEQM: BASEBANDEQMID=3, BASEBANDEQMTYPE=ULDL, UMTSDEMMODE=NULL, SN1=3;""")
    s.mml("2a) Macro SFN sector equipment",
"""ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=0, BaseBandEqmId=0;
ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=1, BaseBandEqmId=1;
ADD EUCELLSECTOREQM: LocalCellId=0, SectorEqmId=2, BaseBandEqmId=2;""")
    s.mml("2b) LampSite sector equipment groups (instead of 2a)",
"""ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=0;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=1;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=2;
ADD EUSECTOREQMGROUP: LocalCellId=0, SectorEqmGroupId=3;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=0, SectorEqmId=0;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=1, SectorEqmId=1;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=2, SectorEqmId=2;
ADD EUSECTOREQMID2GROUP: LocalCellId=0, SectorEqmGroupId=3, SectorEqmId=3;""")
    s.mml("3) Co-processing, cluster, membership, switches",
"""ADD COPROCRES: CoProcResId=0, BaseBandEqmId=3, BundlingClusterType=DMIMO;
ADD DMIMOCLUSTER: DMIMOClusterId=0, CoProcResId=0;
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=0, eNodeBId=0, Mcc="100", Mnc="01";
MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=5;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-1, MuBfAlgoSwitch=MuBfSwitch-1, DMIMOAlgoSwitch=DmimoJTSwitch-1, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-1;
MOD CELLALGOSWITCH: LocalCellId=0, CoordinationAlgoSwitch= MULTI_UE_COORDINATION_OPT_SW-1;
MOD CELLPDCCHALGO: LocalCellId=0, PdcchBfGainOffset=0;
MOD CELLBF: LocalCellId=0, WaitPairingLayerThd=30;""")
    s.para("MAE-Deployment alternative: Feature Configuration Using the MAE-Deployment. OSS GUI may differ from training video.")
    s.para("After this script: go to P6 Activation Verify. If this site is trial inter-eNodeB, continue to P4 before verify.")
    s.footer_end()


def phase4_interenb(book):
    s = book.sheet("P4 Inter-eNodeB Extra", "Phase 4  —  Inter-eNodeB extras (chapter 5) AFTER Phase 3", "TDLEOFD-121501 trial only. Skip this sheet for commercial intra-cell or inter-cell.")
    s.note("TRIAL. MoU required. RGPS-host BBU cannot join an inter-eNodeB cluster (intra-eNodeB D-MIMO only) unless the Cloud BB clock is GPS, not RGPS. Not supported if first- or second-level USU has a backup.", "TRIAL")
    s.figure(f"{FIG}/fig_5_1_inter_enodeb_networking.png", "Figure 5-1  Clock path you must lock before expecting calibration to succeed")
    s.step(4.1, "Confirm Phase 3 downlink D-MIMO is already enabled on the logical SFN cell",
           "Prerequisite function is DmimoJTSwitch. Cluster and co-processing board must sit on the BBU where layer 2 of the SFN cell is configured. Main control = UMPTe/UMPTg/combo.")
    s.step(4.2, "Set TASM.MODE on USUs",
           "Must NOT be FREE for top-level USUs. MAY be FREE for lower-level USUs.")
    s.step(4.3, "Set TASM.SYSCLKSRC",
           "One-level USU: LOCAL on first-level USUs; INTER_SYSCLK on the BBU. Two-level USU: LOCAL on second-level USU; INTER_SYSCLK on first-level USUs; INTER_SYSCLK on the BBU. Details: USU3910-based Multi-BBU Interconnection.")
    s.step(4.4, "Verify lock",
           "DSP SYSCLKSRC: Locked System Clock Source = Interconnection System Clock AND ESN of the NE where the locked source is located is the SAME for all BBUs and USUs.")
    s.step(4.5, "If calibration route search fails",
           "DSP INREC cannot be executed on the secondary BBU to query inter-RRU calibration interference power. In other cases DSP INREC on secondary BBU is allowed.")
    s.footer_end()


def phase5_intercell(book):
    s = book.sheet("P5 Inter-Cell Steps", "Phase 5  —  Inter-cell D-MIMO sequential configuration (6.4.1)", "TDLEOFD-130501. Do NOT use DmimoJTSwitch here — use InterCellDmimoJTSwitch.")
    s.figure(f"{FIG}/fig_6_3_pci_mod3_alignment.png", "Figure 6-3  Complete PCI mod 3 reconstruction BEFORE step 5.8")
    s.step(5.1, "Reconstruct PCI so every cluster member has the same PCI mod 3",
           "If CRS of neighbor faces data REs of the local cell, JT gain is wasted. Plan cluster PCIs AND neighbor PCIs with Huawei PCI planning tool.",
           "Example in source MML: PhyCellId 3, 6, 9 (all mod 3 = 0).")
    s.step(5.2, "Align radio parameters across members",
           "Same bandwidth, frequency, UL-DL subframe assignment, special subframe pattern, CRS port count. Outdoor: common cells only (no SFN mix). LampSite: each pRRU one cell. No simulated-load cells.")
    s.step(5.3, "ADD BASEBANDEQM for LCOP slot and service BBP slot",
           "Example: BASEBANDEQMID=0 on SN1=3 (LCOP), BASEBANDEQMID=1 on SN1=4 (UBBP).")
    s.step(5.4, "ADD CELL / CELLOP / EUCELLSECTOREQM / EUCELLPRIBBEQM for each member",
           "Source example: 8T8R, band 38, DlEarfcn=37850, 20 MHz (N100), SA2, SSP7, MultiRruCellFlag=BOOLEAN_FALSE, CellId 1/2/3.")
    s.step(5.5, "ADD COPROCRES on the dedicated LCOP",
           "WorkMode=COORDINATING_PROCESSING-1&BASEBAND_PROCESSING-0. BundlingClusterType=DMIMO (or ADAPTIVE if mixed). Wrong WorkMode on a non-capable BBP → ALM-26245; services still run on universal baseband only.")
    s.step(5.6, "ADD DMIMOCLUSTER  (NEBF commercial)")
    s.step(5.7, "ADD DMIMOCLUSTERCELL for EVERY member cell (unlike intra-cell’s single cell)",
           "Source example MCC=460 MNC=01 CellId 1/2/3.")
    s.step(5.8, "Per cell: DMSrsPcSinrOffset=5")
    s.step(5.9, "Per cell: SRS auto neighbor measurement",
           "SrsAutoNCellMeasSwitch=ON, NCellSrsMeasA3Offset=-12, NCellSrsTimeMeasSwitch-1.")
    s.step(5.10, "Per cell: turn on BF, MU-BF, InterCellDmimoJTSwitch, quick cal, DL_COORD_SCH_EXP_OPT_SW",
           "This is the actual activation of inter-cell D-MIMO. DL_COORD_SCH_EXP_OPT_SW is recommended when CEU perceived rate is low.")
    s.step(5.11, "(Optional) AvoidCrsConflictInterfSw-1 after PCI mod 3 is aligned")
    s.step(5.12, "Per cell: DmimoA3RsrpOffset=-12 (actual –6 dB), DmimoJtCellRsrpDiffThld=-20 (actual –10 dB)")
    s.step(5.13, "Per cell: RsvResSrsPeriod1=10ms, RsvResSrsPeriod2=20ms — MUST be identical on all cluster/CoMP/AMC members")
    s.step(5.14, "Per cell: BasedA3EdgeUserSwitch-1, EdgeUserA3Offset=-12")
    s.step(5.15, "ADD bidirectional intra-frequency neighbors among all members")
    s.step(5.16, "(Optional) PDCCHAggLvlAdaptStrage=STRATEGYBASEDONCAPACITY")
    s.step(5.17, "CellCqiAdjSchCntThld=50")
    s.step(5.18, "(Optional, LampSite 2T2R) SrsResOptSwitch — HIGH RISK, UBBPe/UBBPd9 only")
    s.footer_end()


def phase5_mml(book):
    s = book.sheet("P5 Inter-Cell MML", "Phase 5  —  Inter-cell MML script (copy/adapt)", "Source 6.4.1.2. Three-cell example. Replace MCC/eNodeBId/PCI/LocalCellId.")
    s.mml("Activation (condensed from source; run in this order)",
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
ADD EUCELLPRIBBEQM:LocalCellId=2,PriBaseBandEqmId=1;
ADD COPROCRES: CoProcResId=0, BaseBandEqmId=0, BundlingClusterType=DMIMO, WorkMode=COORDINATING_PROCESSING-1&BASEBAND_PROCESSING-0;
ADD DMIMOCLUSTER: DMIMOClusterId=0, CoProcResId=0;
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=1, eNodeBId=0, Mcc="460", Mnc="01";
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=2, eNodeBId=0, Mcc="460", Mnc="01";
ADD DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=3, eNodeBId=0, Mcc="460", Mnc="01";
MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=5;
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
    s.para("Next: P6 Activation Verify.")
    s.footer_end()


def phase6_verify(book):
    s = book.sheet("P6 Activation Verify", "Phase 6  —  Activation observation / verification", "Intra-cell = 4.4.2. Inter-cell = 6.4.2. Inter-eNodeB = 4.4.2 + INREC note.")
    s.h2("Intra-cell / inter-eNodeB observation")
    s.step("6A.1", "LST CELL + DSP CELL",
           "Pass if Mode of Multi-RRU Cell = SFN or MPRU_AGGREGATION AND Work Status = Normal for all RRUs AND Cell instance state = Normal.")
    s.step("6A.2", "DSP DMIMOCLUSTERCELL", "Pass if D-MIMO Cluster Cell Status = Normal.")
    s.step("6A.3", "DSP DMIMOCALIBRATION", "Pass if D-MIMO Calibration Result = Succeeded.")
    s.h2("Inter-cell verification (extra DSP CELLCALIBRATION)")
    s.step("6B.1", "DSP CELL + LST CELL", "Cell instance state = Normal AND Work Status = Normal for all RRUs.")
    s.step("6B.2", "DSP CELLCALIBRATION", "Calibration Result = Success.")
    s.step("6B.3", "DSP DMIMOCLUSTERCELL", "DMIMO Cluster Cell Status = Normal.")
    s.step("6B.4", "DSP DMIMOCALIBRATION",
           "Succeeded = SU-JT and MU-JT available. Failed = only non-coherent JT (power overlay, no array gain). Cluster may still be Normal. Intra-cell would have reverted to SFN; inter-cell does NOT revert.")
    s.note("If intra-cell calibration/activation/RRU faults, cells automatically revert to SFN. If inter-cell calibration faults, cells stay in D-MIMO but without coherent JT.", "WARNING")
    s.footer_end()


def phase7_monitor(book):
    s = book.sheet("P7 Network Monitoring", "Phase 7  —  Prove the feature takes effect and watch CEUs", "4.4.3 and 6.4.3. Configure the MAE alarm in the same session.")
    s.h2("Does it take effect?")
    s.table([
        ["Mode", "Primary KPI", "Pass"],
        ["Intra-cell", "L.ChMeas.DMIMO.1Layer.PRB", "> 0"],
        ["Inter-cell", "L.ChMeas.DMIMO.JT.User.Avg", "> 0"],
        ["Both", "L.ChMeas.DMIMO.2Layer.PairPRB … 8Layer.PairPRB", "Non-zero at layers you expect to pair"],
        ["Both", "L.ChMeas.DMIMO.JT.User.Avg / JTUser.RRU.Avg", "JT users present; working RRUs ≥ 2 for true JT"],
        ["Both", "CHR PERIOD_CLUSTER_DMIMO_MR (15 min default)", "ClusterAbnormallatency and CalibrationAbnormallatency stay low"],
    ])
    s.h2("MAE-Access MIMO (Cell) / MIMO (User)")
    s.para("Cell: Enable DMIMO Pairing RB; successful pairing TM7/TM8/TM9 at N layers (N=2..8). Inter-cell also: JT User Num, JT User Work RRU Num, SUJT Schedule RB Num. User: set Test Items = DMIMO (intra-cell) or TDD DMIMO (inter-cell). Inter-cell DMIMO User Indication: 0 non-JT, 1 JT, 255 invalid.")
    s.h2("CEU counters")
    s.para("Intra-cell uses *.BorderUE.JointTransmit / JointReception names. Inter-cell uses *.BorderUE without Joint* suffix (see review workbook 04-5 and 06-5 for full lists).")
    s.h2("STEP 7.1  Create calibration-failure threshold alarm (exact source GUI)")
    s.para("Counter L.CellSectorEqpt.UNA.Dur.Cali. Direction Ascending. Function subset = Network / Measurements Related to Algorithm(LTE) / Cell Sector Algorithm Measurement. For Period=15 min: Threshold=9 (15×80%−3), Offset=3 (15×20%). Alarm when duration ≥ Threshold+Offset; clear when < Threshold−Offset.")
    s.figure(f"{FIG}/mae_step1_threshold_settings.png", "GUI 1  Performance > Threshold Management > Threshold Settings → CellSectorEQUIP Performance → Add")
    s.figure(f"{FIG}/mae_step3_add_threshold_object.png", "GUI 2  Object tab — select NEs")
    s.figure(f"{FIG}/mae_step4_basic_tab.png", "GUI 3  Basic tab — Name CAL Fail, Period 15 minutes")
    s.figure(f"{FIG}/mae_step5_advanced_tab.png", "GUI 4  Advanced tab — counter and 9 / 3")
    s.h2("Expected side effects (do not treat as defects without baseline)")
    s.bullets([
        "RAR success slightly down; access delay slightly up. Inter-cell: RRC setup/reest/HO slightly down; VoLTE/call delay +<10 ms.",
        "Average reported CQI may rise (more TM8) or fall in light-load inter-cell after CRS alignment.",
        "CCE pressure: more DL scheduling per TTI. Watch UL CCE fails.",
        "MU-BF and SFN inter-RRU coordinated-BF counters stop measuring once D-MIMO is the interference-suppression path.",
    ])
    s.footer_end()


def phase8_trouble(book):
    s = book.sheet("P8 Troubleshooting", "Phase 8  —  Possible issues (4.4.4 / 5.4.4 / 6.4.4)", "Work top to bottom. Collect CINR/RSSI before calling Huawei TAC.")
    s.h2("A. DSP DMIMOCALIBRATION not Succeeded")
    s.table([
        ["Result", "Meaning", "Do this"],
        ["Exception occurred in internal calibration", "Intra-RRU cal failed", "DSP CELLCALIBRATION → identify RRU → Huawei TAC"],
        ["Route search failed", "No path; external interference", "Eliminate interference. Inter-eNB: DSP INREC may be blocked on secondary BBU."],
        ["Reciprocity calibration failed", "Route OK, thresholds missed", "Record Calibration Signal CINR and RSSI → Huawei TAC"],
    ])
    s.h2("B. DSP DMIMOCLUSTERCELL not Normal")
    s.table([
        ["Status", "Do this"],
        ["Incorrect configurations or unavailable licenses / Clock exceptions (inter-eNB)", "Switch, RRU count, cluster, license. Inter-cell also: CRS ports, TDD SRS mode, reserved SRS 1/2, subframe config consistent. Inter-eNB: inter-eNB link + DSP SYSCLKSRC ESN match."],
        ["Limited hardware capacity", "Intra: UBBPd/UBBPe. Inter-cell: same UBBPd9 or UBBPe"],
        ["Route application failures", "Routing bandwidth"],
        ["Cell abnormal", "DSP CELL"],
        ["Insufficient / abnormal co-processing resources", "Intra: 4.3.3 board rules. Inter-cell: resources on LCOP"],
        ["Channel calibration failures", "Section A"],
        ["Cluster ID conflicts", "mod 6 / mod 24 / uniqueness / even ID"],
        ["Clusters being established", "Wait"],
        ["LampSite minimum cells / mixed eNodeB types / micro TM8 dual-layer", "Inter-cell only — see 06-5 in the review workbook"],
    ])
    s.h2("C. Alarms")
    s.table([
        ["Alarm", "Typical cause"],
        ["ALM-26245 Configuration Data Inconsistency", "COORDINATING_PROCESSING on a BBP that cannot do it; or LCOP/BBP cannot central-control inter-cell — replace BBP per 6.3.3"],
        ["ALM-26203 Board Software Program Error", "Same BBP central-control check"],
        ["ALM-29243 Cell Capability Degraded / ALM-29240 Cell Unavailable", "SrsResOptSwitch on UBBPd4/LBBPd"],
        ["User-defined CAL Fail threshold", "L.CellSectorEqpt.UNA.Dur.Cali ≥ 9+3 in 15-min example"],
    ])
    s.h2("D. Dual-stream BF block errors / disconnects")
    s.para("When DmimoJTSwitch is selected, consecutive block errors may arise during dual-stream beamforming, even leading to network disconnections. Workarounds are in Beamforming (TDD), not in this D-MIMO document.")
    s.footer_end()


def phase9_rollback(book):
    s = book.sheet("P9 Deactivation Rollback", "Phase 9  —  Deactivation (source command examples)", "Restore other parameters to the live-network baseline, not blindly to these examples.")
    s.h2("Intra-cell rollback order")
    s.mml("Intra-cell deactivation",
"""MOD CELLPCALGO: LocalCellId=0, DMSrsPcSinrOffset=0;
MOD CELLALGOSWITCH: LocalCellId=0, BfAlgoSwitch=BfSwitch-0, MuBfAlgoSwitch=MuBfSwitch-0, DMIMOAlgoSwitch=DmimoJTSwitch-0, EnhChnCalSwitch=QUICK_CHN_CAL_SWITCH-0;
MOD CELLALGOSWITCH: LocalCellId=0, CoordinationAlgoSwitch= MULTI_UE_COORDINATION_OPT_SW-0;
RMV DMIMOCLUSTERCELL: DMIMOClusterId=0, CellId=0, eNodeBId=0, Mcc="100", Mnc="01";
RMV DMIMOCLUSTER: DMIMOClusterId=0;
MOD CELLPDCCHALGO: LocalCellId=0, PdcchBfGainOffset=-127;
MOD CELLBF: LocalCellId=0, WaitPairingLayerThd=0;""")
    s.h2("Inter-cell rollback order")
    s.mml("Inter-cell deactivation (adapt CellId/MCC to what you added)",
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
    s.note("Source activation used CellId 1/2/3 and MCC 460; deactivation sample uses CellId 0/1/2 and MCC 100. Always RMV the exact tuple you ADDed.", "NOTICE")
    s.para("Inter-eNodeB rollback: first disable inter-eNodeB clock extras if they were added only for this trial, then follow intra-cell rollback. Do not disturb a GPS/RGPS plan that other services still need.")
    s.footer_end()


def checklist(book):
    s = book.sheet("CL Master Checklist", "Master deployment checklist (one row per cluster)", "Print this sheet as the site pack cover. Fill Status as ☐ / ☑ / N/A.")
    s.table([
        ["#", "Gate", "Evidence", "Owner", "Status"],
        ["0.1", "eRAN22.1 TDD confirmed", "LST SYS screenshot", "RAN", "☐"],
        ["0.3", "If 121501: MoU + risk acceptance", "MoU file", "PM", "☐"],
        ["0.4", "Parent license sufficient", "LST LICENSE", "RAN", "☐"],
        ["0.5", "Cluster ID + (inter-cell) PCI plan from Huawei tools", "Planning sheet", "Huawei/NP", "☐"],
        ["1.1", "Mode selected (111505 / 121501 / 130501)", "This pack Cover table", "NP", "☐"],
        ["1.1b", "ISD / overlap / LOS / speed / UE count fit", "DT or planning plot", "NP", "☐"],
        ["2.2", "Prerequisites ON (SFN/BF/MU-BF as required)", "LST CELLALGOSWITCH", "RAN", "☐"],
        ["2.3", "Exclusive functions OFF", "Switch dump", "RAN", "☐"],
        ["2.4", "Energy-saving/sleep functions OFF on cluster cells", "Shutdown/LowPwr/DRX dump", "RAN", "☐"],
        ["2.5", "BBP / co-proc or LCOP / RF / bandwidth / band / frame offset", "Inventory vs 4.3.3 / 6.3.3", "Field", "☐"],
        ["2.6", "Cluster ID uniqueness, mod 6, mod 24, even-ID LampSite", "Planning vs LST DMIMOCLUSTER", "NP", "☐"],
        ["3.x / 5.x", "MML/MAE executed in documented order", "Script log", "RAN", "☐"],
        ["4.x", "If inter-eNB: clock MODE/SYSCLKSRC/ESN match, no USU backup, RGPS-host not in cluster", "DSP SYSCLKSRC", "RAN", "☐"],
        ["6", "DSP CELL / DMIMOCLUSTERCELL / DMIMOCALIBRATION (+ CELLCALIBRATION inter-cell) pass", "Command output saved", "RAN", "☐"],
        ["7", "Pairing counter > 0; MAE CAL Fail alarm created (Thd 9 / Off 3 @ 15 min)", "KPI + threshold ID", "NOC", "☐"],
        ["7b", "CEU DL/UL counters baselined 3–7 days", "Pre/post report", "Opt", "☐"],
        ["8", "No open 26245/26203/29240/29243; cal latency acceptable", "Alarm list", "NOC", "☐"],
        ["Sign-off", "Cluster accepted / rolled back", "Change ticket", "RAN Lead", "☐"],
    ])
    s.h2("Recommended commercial values (quick reference)")
    s.table([
        ["Parameter", "Intra-cell", "Inter-cell"],
        ["BfWeightNormalizeMode", "NEBF", "NEBF"],
        ["DMIMOAlgoSwitch", "DmimoJTSwitch", "InterCellDmimoJTSwitch"],
        ["EnhChnCalSwitch", "QUICK_CHN_CAL_SWITCH", "QUICK_CHN_CAL_SWITCH"],
        ["DMSrsPcSinrOffset", "5", "5"],
        ["Coordination / DlSchExt", "MULTI_UE_COORDINATION_OPT_SW", "DL_COORD_SCH_EXP_OPT_SW"],
        ["AvoidCrsConflictInterfSw", "—", "ON after PCI mod 3 aligned"],
        ["DmimoA3RsrpOffset / JtCellRsrpDiffThld", "—", "-12 / -20"],
        ["RsvResSrsPeriod1/2", "—", "10ms / 20ms (identical in cluster)"],
        ["WaitPairingLayerThd", "30 if heavy+multi-layer else 0", "—"],
        ["PdcchBfGainOffset / Agg strategy", "0 if coverage-based strategy", "STRATEGYBASEDONCAPACITY"],
        ["TmAccelerationSwitch", "INITIAL_ACCESS_TO_BF", "—"],
        ["CellCqiAdjSchCntThld", "—", "50"],
        ["CoProc WorkMode", "as needed", "COORDINATING_PROCESSING only (no baseband)"],
    ])
    s.footer_end()


if __name__ == "__main__":
    build()
