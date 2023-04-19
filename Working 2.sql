DECLARE @ReportingYear as numeric(18,0)
DECLARE @CurrentDate as date
SET @ReportingYear = 2022;
SET @CurrentDate = GETDATE();

--Temp Table for rows with a MY Program ID Event with Single program
WITH
[cteMYProgIDSingleProgram] AS
(
SELECT DISTINCT (SA1.MYProgramID) AS SA1MYProgramID, Count(SA1.Year) AS 'MYProgramID Count'
			, Min(SA1.Year) AS 'Min Impact Year'
			, Max(SA1.Year) AS 'Max Impact Year'
			, SUBSTRING((SELECT (','+Rtrim(a1.Year)) 
						from [dbo].[tb_StrategicAssessments] a1
						LEFT JOIN [tb_QuantitativeAnalyses] a2 ON a2.SADOCID = a1.SADOCID
						LEFT JOIN [vw_LegalContracts] a3 ON a3.SADOCID = a1.SADOCID
						LEFT JOIN tb_VSSMFinancialChecklists a4 on a4.SADOCID = a1.SADOCID
						LEFT JOIN [tb_PropertyEquityAssessment] a5 ON a5.SADOCID=a1.SADOCID
						WHERE a1.MYProgramID = SA1.MYProgramID
						AND a1.Year >= 2016 AND a1.TestProg IS NULL 
						AND a4.FirstApprovalDate IS NOT NULL 
						AND a1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
						AND a3.FormTitle <> 'Legal Amendment'
						AND a1.Multi = 'Yes'
						FOR XML PATH('')),2,1000) AS 'Impact Years'
			FROM [dbo].[tb_StrategicAssessments] SA1
			LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA1 ON QA1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].vw_LegalContracts LC1 ON LC1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC1 on FC1.SADOCID = SA1.SADOCID
			LEFT JOIN [tb_PropertyEquityAssessment] PE1 on PE1.SADOCID = SA1.SADOCID
			WHERE	SA1.Year >= 2016 AND SA1.TestProg IS NULL 
				AND FC1.FirstApprovalDate IS NOT NULL 
				AND SA1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
				AND LC1.FormTitle <> 'Legal Amendment'
				AND SA1.Multi = 'Yes'
		 GROUP BY SA1.MYProgramID
		 HAVING Count(SA1.Year) = 1),

--Temp Table for rows with a MY Program ID Event with multiple programs in different and same Impact Years
[cteMYMultipleProgramSYDY] AS
(
SELECT (SA1.MYProgramID) AS SA1MYProgramID, Count(SA1.Year) AS 'MYProgramID Count'
			, Min(SA1.Year) AS 'Min Impact Year'
			, Max(SA1.Year) AS 'Max Impact Year'
			, SUBSTRING((SELECT (','+Rtrim(a1.Year)) 
						from [dbo].[tb_StrategicAssessments] a1
						LEFT JOIN [tb_QuantitativeAnalyses] a2 ON a2.SADOCID = a1.SADOCID
						LEFT JOIN [vw_LegalContracts] a3 ON a3.SADOCID = a1.SADOCID
						LEFT JOIN tb_VSSMFinancialChecklists a4 on a4.SADOCID = a1.SADOCID
						LEFT JOIN [tb_PropertyEquityAssessment] a5 on a5.SADOCID = a1.SADOCID
						WHERE a1.MYProgramID = SA1.MYProgramID
						AND a1.Year >= 2016 AND a1.TestProg IS NULL 
						AND a4.FirstApprovalDate IS NOT NULL 
						AND a1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
						AND a3.FormTitle <> 'Legal Amendment'
						AND a1.Multi = 'Yes'
						FOR XML PATH('')),2,1000) AS 'Impact Years'
			FROM [dbo].[tb_StrategicAssessments] SA1
			LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA1 ON QA1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].vw_LegalContracts LC1 ON LC1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC1 on FC1.SADOCID = SA1.SADOCID
			LEFT JOIN [tb_PropertyEquityAssessment] PE1 on PE1.SADOCID = SA1.SADOCID
			WHERE	SA1.Year >= 2016 AND SA1.TestProg IS NULL 
				AND FC1.FirstApprovalDate IS NOT NULL 
				AND SA1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
				AND LC1.FormTitle <> 'Legal Amendment'
				AND SA1.Multi = 'Yes'
		 GROUP BY SA1.MYProgramID
		 HAVING Count(SA1.Year) > 1),

--Temp Table for rows with a MY Program ID Event with multiple programs in same Impact Years
[cteMYMultipleProgramSameYear] AS
(
SELECT (SA1.MYProgramID) AS SA1MYProgramID,SA1.Year, Count(SA1.Year) AS 'MYProgramID Count'
			, CAST(Min(SA1.EventStart) AS DATE) AS 'Min Event Startdate'
			, CAST(Max(SA1.EventStart) AS DATE) AS 'Max Event Startdate'
			, SUBSTRING((SELECT (','+Rtrim(a1.EventStart)) 
						from [dbo].[tb_StrategicAssessments] a1
						LEFT JOIN [tb_QuantitativeAnalyses] a2 ON a2.SADOCID = a1.SADOCID
						LEFT JOIN [vw_LegalContracts] a3 ON a3.SADOCID = a1.SADOCID
						LEFT JOIN tb_VSSMFinancialChecklists a4 on a4.SADOCID = a1.SADOCID
						LEFT JOIN [tb_PropertyEquityAssessment] a5 on a5.SADOCID = a1.SADOCID
						WHERE a1.MYProgramID = SA1.MYProgramID
						AND a1.Year >= 2016 AND a1.TestProg IS NULL 
						AND a4.FirstApprovalDate IS NOT NULL 
						AND a1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
						AND a3.FormTitle <> 'Legal Amendment'
						AND a1.Multi = 'Yes'
				FOR XML PATH('')),2,1000) AS 'Event Start Dates'
			,CASE WHEN ((YEAR(Min(SA1.EventStart)) = YEAR(@CurrentDate) AND YEAR(Max(SA1.EventStart)) <> YEAR(@CurrentDate)) OR 
						(YEAR(Min(SA1.EventStart)) <> YEAR(@CurrentDate) AND YEAR(Max(SA1.EventStart)) = YEAR(@CurrentDate))) THEN 'Different Year'
			ELSE 'Same year' END AS SameYear
			,CASE WHEN YEAR(Min(SA1.EventStart)) = YEAR(@CurrentDate) THEN Min(SA1.EventStart)
				  WHEN YEAR(Max(SA1.EventStart)) = YEAR(@CurrentDate) THEN Max(SA1.EventStart) END AS 'CYEventDate'
			FROM [dbo].[tb_StrategicAssessments] SA1
			LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA1 ON QA1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].vw_LegalContracts LC1 ON LC1.SADOCID = SA1.SADOCID
			LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC1 on FC1.SADOCID = SA1.SADOCID
			LEFT JOIN [tb_PropertyEquityAssessment] PE1 on PE1.SADOCID = SA1.SADOCID
			WHERE	SA1.Year >= 2016 AND SA1.TestProg IS NULL 
				AND FC1.FirstApprovalDate IS NOT NULL 
				AND SA1.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
				AND LC1.FormTitle <> 'Legal Amendment'
				AND SA1.Multi = 'Yes'
				AND SA1.Year = @ReportingYear
		 GROUP BY SA1.MYProgramID,SA1.Year
		 HAVING COUNT(SA1.Year) > 1 
),

[cteFinalScript] AS
(
-- UNION 1 - ALL ROWS for Non Multi year Programs (A.K.A. - Single Year Programs)
SELECT DISTINCT SA.SADOCID, 
SA.RoutingNum AS 'Routing Number',
SA.InitBy AS 'Initiated By',
SA.Channel AS 'Channel',
SA.Region AS 'Region/Client Classification',
SA.Division AS 'Division',
SA.ProposalTitle AS 'Program Name',
SA.MYProgramID AS 'Multi-Year Prog ID',
SA.ExecutiveSummary AS 'Executive Summary',
SA.Status AS 'Status',
SA.SelectType AS 'Promotion Type',
SA.Category AS 'Category',
SA.SubCategory AS 'Subcategory',
SA.Multi AS 'Multi-Year Agreement',
SA.EventStart AS 'Event Start Date',
SA.EventEnd AS 'Event End Date',
SA.TermStart AS 'Multi-Year Term Start Date',
SA.TermEnd AS 'Multi-Year Term End Date',
SA.PromoRBM AS 'GM Contact',
SA.RWContact AS 'Jack Morton Contact',
SA.RWPD AS 'Jack Morton Director',
SA.Company AS 'Promoter Company Name',
SA.Add1 AS 'Promoter Address',
SA.City AS 'Promoter City',
SA.State AS 'Promoter State',
SA.Zip AS 'Promoter Zip',
SA.PromoName AS 'Promoter Contact Name',
SA.EMail AS 'Promoter Email',
SA.Phone AS 'Promoter Phone',
SA.Year AS 'Impact Year',
SA.TestProg AS 'TestProg',
SA.CurrentContractYear AS 'Contract Year',
QA.SFCost AS 'Negotiated Sponsorship Fee',
QA.TotalVehCost1_2 AS 'Courtesy Vehicles Cost',
QA.SDVCost AS 'Sweepstakes/Donation Vehicles Cost',
QA.AICost AS 'Additional Contractual Cost',
QA.SCostSum AS 'Total Contractual Cost',
QA.TotalVehCost1_1_1 AS 'Display Vehicles Cost',
QA.TotalVehCostRD_1 AS 'Non-Contractual Loaned Vehicles Cost',
QA.EstCost_1 AS 'Estimated Activation Cost (w/out vehicles)',
QA.CostSum AS 'Total Estimated Activation Cost (with vehicles)',
QA.ACostSum AS 'Total Estimated Program Cost',
QA.MultiCost1 AS 'Negotiated Multi-Year Sponsorship Fee',
QA.MultiCost3 AS 'Est. Multi-Year Courtesy Vehicle Cost',
QA.MultiCost7 AS 'Est. Multi-Year Sweeps/Donation Vehicle Cost',
QA.MultiCost2 AS 'Negotiated Multi-Year Additional Required Cost:',
QA.MultiCost8 AS 'Multi-Year Total Contractual Cost',
QA.MultiCost4 AS 'Est. Multi-Year Total Activation (incl. vehicles)',
QA.MultiCost9 AS 'Multi-Year Total Est. Program Cost',
QA.MultiNotes AS 'Multi-Year Cost Notes',
QA.TotalPBR AS 'Total GM Asset Value',
QA.peov AS 'Equity & Opportunity Value',
QA.TotalProgValue AS 'Total Program Value',
QA.SysParmQ3 AS 'GM Value = or > Contractual Cost?',
LC.Outclause AS 'If multi-year - is there an out-clause?',
LC.FormTitle AS 'Contract Type',
LC.MultiDate1 AS 'Out Clause Notification Date',
LC.PaymentTerms AS 'Payment Terms',
LC.PayTermDetail AS 'Payment Terms Details',
LC.ActualContractStartDate AS 'Contract Execution Date',
LC.ActualContractEndDate AS 'Contract Termination Date',
FC.FirstApprovalDate AS 'Date Financial Checklist First Approved',
FC.StatusDate AS 'Date of Final Financial Approval',
FC.RFI AS 'Finance Budget Manager',
QA.EstActDays AS 'Estimated Activation Days',
PE.TotScreenScore AS 'PEA Screening Score',
QA.PEWeightScore AS 'Property Equity Weight',
QA.MarketWeightScore AS 'Market Weight',
QA.AudWeightScore AS 'Audience Weight',
QA.PEWeightScore+QA.MarketWeightScore+QA.AudWeightScore AS 'Total Weight',
QA.peov AS 'Property Equity & Opportunity Value',
SA.RenewalID AS 'RenewalID',
'Union - 1' AS 'Union',
@ReportingYear AS ReportingYear
FROM [dbo].[tb_StrategicAssessments] SA 
LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA ON QA.SADOCID = SA.SADOCID
LEFT JOIN [dbo].vw_LegalContracts LC ON LC.SADOCID = SA.SADOCID
LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC on FC.SADOCID = SA.SADOCID
LEFT JOIN [tb_PropertyEquityAssessment] PE on PE.SADOCID = SA.SADOCID
WHERE	SA.Year >= 2016 AND SA.TestProg IS NULL 
	AND FC.FirstApprovalDate IS NOT NULL 
	AND SA.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
	AND LC.FormTitle <> 'Legal Amendment'
	AND SA.Multi = 'NO' 

UNION

--UNION 2 -  ALL Rows for Multi Year Prog ID with Program in Single year
SELECT DISTINCT SA.SADOCID,
SA.RoutingNum AS 'Routing Number',
SA.InitBy AS 'Initiated By',
SA.Channel AS 'Channel',
SA.Region AS 'Region/Client Classification',
SA.Division AS 'Division',
SA.ProposalTitle AS 'Program Name',
SA.MYProgramID AS 'Multi-Year Prog ID',
SA.ExecutiveSummary AS 'Executive Summary',
SA.Status AS 'Status',
SA.SelectType AS 'Promotion Type',
SA.Category AS 'Category',
SA.SubCategory AS 'Subcategory',
SA.Multi AS 'Multi-Year Agreement',
SA.EventStart AS 'Event Start Date',
SA.EventEnd AS 'Event End Date',
SA.TermStart AS 'Multi-Year Term Start Date',
SA.TermEnd AS 'Multi-Year Term End Date',
SA.PromoRBM AS 'GM Contact',
SA.RWContact AS 'Jack Morton Contact',
SA.RWPD AS 'Jack Morton Director',
SA.Company AS 'Promoter Company Name',
SA.Add1 AS 'Promoter Address',
SA.City AS 'Promoter City',
SA.State AS 'Promoter State',
SA.Zip AS 'Promoter Zip',
SA.PromoName AS 'Promoter Contact Name',
SA.EMail AS 'Promoter Email',
SA.Phone AS 'Promoter Phone',
SA.Year AS 'Impact Year',
SA.TestProg AS 'TestProg',
SA.CurrentContractYear AS 'Contract Year',
QA.SFCost AS 'Negotiated Sponsorship Fee',
QA.TotalVehCost1_2 AS 'Courtesy Vehicles Cost',
QA.SDVCost AS 'Sweepstakes/Donation Vehicles Cost',
QA.AICost AS 'Additional Contractual Cost',
QA.SCostSum AS 'Total Contractual Cost',
QA.TotalVehCost1_1_1 AS 'Display Vehicles Cost',
QA.TotalVehCostRD_1 AS 'Non-Contractual Loaned Vehicles Cost',
QA.EstCost_1 AS 'Estimated Activation Cost (w/out vehicles)',
QA.CostSum AS 'Total Estimated Activation Cost (with vehicles)',
QA.ACostSum AS 'Total Estimated Program Cost',
QA.MultiCost1 AS 'Negotiated Multi-Year Sponsorship Fee',
QA.MultiCost3 AS 'Est. Multi-Year Courtesy Vehicle Cost',
QA.MultiCost7 AS 'Est. Multi-Year Sweeps/Donation Vehicle Cost',
QA.MultiCost2 AS 'Negotiated Multi-Year Additional Required Cost:',
QA.MultiCost8 AS 'Multi-Year Total Contractual Cost',
QA.MultiCost4 AS 'Est. Multi-Year Total Activation (incl. vehicles)',
QA.MultiCost9 AS 'Multi-Year Total Est. Program Cost',
QA.MultiNotes AS 'Multi-Year Cost Notes',
QA.TotalPBR AS 'Total GM Asset Value',
QA.peov AS 'Equity & Opportunity Value',
QA.TotalProgValue AS 'Total Program Value',
QA.SysParmQ3 AS 'GM Value = or > Contractual Cost?',
LC.Outclause AS 'If multi-year - is there an out-clause?',
LC.FormTitle AS 'Contract Type',
LC.MultiDate1 AS 'Out Clause Notification Date',
LC.PaymentTerms AS 'Payment Terms',
LC.PayTermDetail AS 'Payment Terms Details',
LC.ActualContractStartDate AS 'Contract Execution Date',
LC.ActualContractEndDate AS 'Contract Termination Date',
FC.FirstApprovalDate AS 'Date Financial Checklist First Approved',
FC.StatusDate AS 'Date of Final Financial Approval',
FC.RFI AS 'Finance Budget Manager',
QA.EstActDays AS 'Estimated Activation Days',
PE.TotScreenScore AS 'PEA Screening Score',
QA.PEWeightScore AS 'Property Equity Weight',
QA.MarketWeightScore AS 'Market Weight',
QA.AudWeightScore AS 'Audience Weight',
QA.PEWeightScore+QA.MarketWeightScore+QA.AudWeightScore AS 'Total Weight',
QA.peov AS 'Property Equity & Opportunity Value',
SA.RenewalID AS 'RenewalID',
'Union - 2' AS 'Union',
@ReportingYear AS ReportingYear
FROM [dbo].[tb_StrategicAssessments] SA 
LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA ON QA.SADOCID = SA.SADOCID
LEFT JOIN [dbo].vw_LegalContracts LC ON LC.SADOCID = SA.SADOCID
LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC on FC.SADOCID = SA.SADOCID
LEFT JOIN [tb_PropertyEquityAssessment] PE on PE.SADOCID = SA.SADOCID
INNER JOIN [cteMYProgIDSingleProgram] IY ON IY.[SA1MYProgramID]=SA.MYProgramID AND IY.[Impact Years]=SA.Year
WHERE	SA.Year >= 2016 AND SA.TestProg IS NULL 
	AND FC.FirstApprovalDate IS NOT NULL 
	AND SA.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
	AND LC.FormTitle <> 'Legal Amendment'
	AND SA.Multi = 'YES'

UNION

--UNION 3 -  ALL Rows with more than one MULTI YEAR Program ID, If IMPACT YEAR IN REPORTING YEAR then Reporting year; ELSE LATEST YEAR
SELECT DISTINCT SA.SADOCID,
SA.RoutingNum AS 'Routing Number',
SA.InitBy AS 'Initiated By',
SA.Channel AS 'Channel',
SA.Region AS 'Region/Client Classification',
SA.Division AS 'Division',
SA.ProposalTitle AS 'Program Name',
SA.MYProgramID AS 'Multi-Year Prog ID',
SA.ExecutiveSummary AS 'Executive Summary',
SA.Status AS 'Status',
SA.SelectType AS 'Promotion Type',
SA.Category AS 'Category',
SA.SubCategory AS 'Subcategory',
SA.Multi AS 'Multi-Year Agreement',
SA.EventStart AS 'Event Start Date',
SA.EventEnd AS 'Event End Date',
SA.TermStart AS 'Multi-Year Term Start Date',
SA.TermEnd AS 'Multi-Year Term End Date',
SA.PromoRBM AS 'GM Contact',
SA.RWContact AS 'Jack Morton Contact',
SA.RWPD AS 'Jack Morton Director',
SA.Company AS 'Promoter Company Name',
SA.Add1 AS 'Promoter Address',
SA.City AS 'Promoter City',
SA.State AS 'Promoter State',
SA.Zip AS 'Promoter Zip',
SA.PromoName AS 'Promoter Contact Name',
SA.EMail AS 'Promoter Email',
SA.Phone AS 'Promoter Phone',
SA.Year AS 'Impact Year',
SA.TestProg AS 'TestProg',
SA.CurrentContractYear AS 'Contract Year',
QA.SFCost AS 'Negotiated Sponsorship Fee',
QA.TotalVehCost1_2 AS 'Courtesy Vehicles Cost',
QA.SDVCost AS 'Sweepstakes/Donation Vehicles Cost',
QA.AICost AS 'Additional Contractual Cost',
QA.SCostSum AS 'Total Contractual Cost',
QA.TotalVehCost1_1_1 AS 'Display Vehicles Cost',
QA.TotalVehCostRD_1 AS 'Non-Contractual Loaned Vehicles Cost',
QA.EstCost_1 AS 'Estimated Activation Cost (w/out vehicles)',
QA.CostSum AS 'Total Estimated Activation Cost (with vehicles)',
QA.ACostSum AS 'Total Estimated Program Cost',
QA.MultiCost1 AS 'Negotiated Multi-Year Sponsorship Fee',
QA.MultiCost3 AS 'Est. Multi-Year Courtesy Vehicle Cost',
QA.MultiCost7 AS 'Est. Multi-Year Sweeps/Donation Vehicle Cost',
QA.MultiCost2 AS 'Negotiated Multi-Year Additional Required Cost:',
QA.MultiCost8 AS 'Multi-Year Total Contractual Cost',
QA.MultiCost4 AS 'Est. Multi-Year Total Activation (incl. vehicles)',
QA.MultiCost9 AS 'Multi-Year Total Est. Program Cost',
QA.MultiNotes AS 'Multi-Year Cost Notes',
QA.TotalPBR AS 'Total GM Asset Value',
QA.peov AS 'Equity & Opportunity Value',
QA.TotalProgValue AS 'Total Program Value',
QA.SysParmQ3 AS 'GM Value = or > Contractual Cost?',
LC.Outclause AS 'If multi-year - is there an out-clause?',
LC.FormTitle AS 'Contract Type',
LC.MultiDate1 AS 'Out Clause Notification Date',
LC.PaymentTerms AS 'Payment Terms',
LC.PayTermDetail AS 'Payment Terms Details',
LC.ActualContractStartDate AS 'Contract Execution Date',
LC.ActualContractEndDate AS 'Contract Termination Date',
FC.FirstApprovalDate AS 'Date Financial Checklist First Approved',
FC.StatusDate AS 'Date of Final Financial Approval',
FC.RFI AS 'Finance Budget Manager',
QA.EstActDays AS 'Estimated Activation Days',
PE.TotScreenScore AS 'PEA Screening Score',
QA.PEWeightScore AS 'Property Equity Weight',
QA.MarketWeightScore AS 'Market Weight',
QA.AudWeightScore AS 'Audience Weight',
QA.PEWeightScore+QA.MarketWeightScore+QA.AudWeightScore AS 'Total Weight',
QA.peov AS 'Property Equity & Opportunity Value',
SA.RenewalID AS 'RenewalID',
'Union - 3' AS 'Union',
@ReportingYear AS ReportingYear
FROM [dbo].[tb_StrategicAssessments] SA 
LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA ON QA.SADOCID = SA.SADOCID
LEFT JOIN [dbo].vw_LegalContracts LC ON LC.SADOCID = SA.SADOCID
LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC on FC.SADOCID = SA.SADOCID
LEFT JOIN [tb_PropertyEquityAssessment] PE on PE.SADOCID = SA.SADOCID
LEFT JOIN [cteMYProgIDSingleProgram] MYSP ON MYSP.SA1MYProgramID=SA.MYProgramID
INNER JOIN (select U1.SA1MYProgramID, U1.[MYProgramID Count],
				U1.[Impact Years], U1.[Min Impact Year],U1.[Max Impact Year],
				CASE	 WHEN  [Impact Years]  Like CONCAT('%',@ReportingYear,'%') then @ReportingYear
				WHEN U1.[Min Impact Year]>=@ReportingYear then U1.[Min Impact Year]
				WHEN U1.[Max Impact Year]<@ReportingYear then U1.[Max Impact Year]
				ELSE U1.[Min Impact Year] END AS SA1EligibleYear
				from [cteMYMultipleProgramSYDY] U1
				LEFT JOIN cteMYMultipleProgramSameYear U2 
				ON U1.SA1MYProgramID=U2.SA1MYProgramID
				WHERE U2.SA1MYProgramID IS NULL) SYDY 
			ON SA.MYProgramID = SYDY.SA1MYProgramID AND SA.Year=SYDY.[SA1EligibleYear]
WHERE	SA.Year >= 2016 AND SA.TestProg IS NULL
	AND FC.FirstApprovalDate IS NOT NULL
	AND SA.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
	AND LC.FormTitle <> 'Legal Amendment'
	AND SA.Multi = 'YES'
	AND MYSP.SA1MYProgramID IS NULL

UNION

--UNION 4 -  ALL Row with more than one MY Program ID over the years and with more than MULTI YEAR Program ID in the same year, WHen there are More than one Entry of the same year
SELECT DISTINCT SA.SADOCID,
SA.RoutingNum AS 'Routing Number',
SA.InitBy AS 'Initiated By',
SA.Channel AS 'Channel',
SA.Region AS 'Region/Client Classification',
SA.Division AS 'Division',
SA.ProposalTitle AS 'Program Name',
SA.MYProgramID AS 'Multi-Year Prog ID',
SA.ExecutiveSummary AS 'Executive Summary',
SA.Status AS 'Status',
SA.SelectType AS 'Promotion Type',
SA.Category AS 'Category',
SA.SubCategory AS 'Subcategory',
SA.Multi AS 'Multi-Year Agreement',
SA.EventStart AS 'Event Start Date',
SA.EventEnd AS 'Event End Date',
SA.TermStart AS 'Multi-Year Term Start Date',
SA.TermEnd AS 'Multi-Year Term End Date',
SA.PromoRBM AS 'GM Contact',
SA.RWContact AS 'Jack Morton Contact',
SA.RWPD AS 'Jack Morton Director',
SA.Company AS 'Promoter Company Name',
SA.Add1 AS 'Promoter Address',
SA.City AS 'Promoter City',
SA.State AS 'Promoter State',
SA.Zip AS 'Promoter Zip',
SA.PromoName AS 'Promoter Contact Name',
SA.EMail AS 'Promoter Email',
SA.Phone AS 'Promoter Phone',
SA.Year AS 'Impact Year',
SA.TestProg AS 'TestProg',
SA.CurrentContractYear AS 'Contract Year',
QA.SFCost AS 'Negotiated Sponsorship Fee',
QA.TotalVehCost1_2 AS 'Courtesy Vehicles Cost',
QA.SDVCost AS 'Sweepstakes/Donation Vehicles Cost',
QA.AICost AS 'Additional Contractual Cost',
QA.SCostSum AS 'Total Contractual Cost',
QA.TotalVehCost1_1_1 AS 'Display Vehicles Cost',
QA.TotalVehCostRD_1 AS 'Non-Contractual Loaned Vehicles Cost',
QA.EstCost_1 AS 'Estimated Activation Cost (w/out vehicles)',
QA.CostSum AS 'Total Estimated Activation Cost (with vehicles)',
QA.ACostSum AS 'Total Estimated Program Cost',
QA.MultiCost1 AS 'Negotiated Multi-Year Sponsorship Fee',
QA.MultiCost3 AS 'Est. Multi-Year Courtesy Vehicle Cost',
QA.MultiCost7 AS 'Est. Multi-Year Sweeps/Donation Vehicle Cost',
QA.MultiCost2 AS 'Negotiated Multi-Year Additional Required Cost:',
QA.MultiCost8 AS 'Multi-Year Total Contractual Cost',
QA.MultiCost4 AS 'Est. Multi-Year Total Activation (incl. vehicles)',
QA.MultiCost9 AS 'Multi-Year Total Est. Program Cost',
QA.MultiNotes AS 'Multi-Year Cost Notes',
QA.TotalPBR AS 'Total GM Asset Value',
QA.peov AS 'Equity & Opportunity Value',
QA.TotalProgValue AS 'Total Program Value',
QA.SysParmQ3 AS 'GM Value = or > Contractual Cost?',
LC.Outclause AS 'If multi-year - is there an out-clause?',
LC.FormTitle AS 'Contract Type',
LC.MultiDate1 AS 'Out Clause Notification Date',
LC.PaymentTerms AS 'Payment Terms',
LC.PayTermDetail AS 'Payment Terms Details',
LC.ActualContractStartDate AS 'Contract Execution Date',
LC.ActualContractEndDate AS 'Contract Termination Date',
FC.FirstApprovalDate AS 'Date Financial Checklist First Approved',
FC.StatusDate AS 'Date of Final Financial Approval',
FC.RFI AS 'Finance Budget Manager',
QA.EstActDays AS 'Estimated Activation Days',
PE.TotScreenScore AS 'PEA Screening Score',
QA.PEWeightScore AS 'Property Equity Weight',
QA.MarketWeightScore AS 'Market Weight',
QA.AudWeightScore AS 'Audience Weight',
QA.PEWeightScore+QA.MarketWeightScore+QA.AudWeightScore AS 'Total Weight',
QA.peov AS 'Property Equity & Opportunity Value',
SA.RenewalID AS 'RenewalID',
'Union - 4' AS 'Union',
@ReportingYear AS ReportingYear
FROM [dbo].[tb_StrategicAssessments] SA 
LEFT JOIN [dbo].[tb_QuantitativeAnalyses] QA ON QA.SADOCID = SA.SADOCID
LEFT JOIN [dbo].vw_LegalContracts LC ON LC.SADOCID = SA.SADOCID
LEFT JOIN [dbo].tb_VSSMFinancialChecklists FC on FC.SADOCID = SA.SADOCID
LEFT JOIN [tb_PropertyEquityAssessment] PE on PE.SADOCID = SA.SADOCID
INNER JOIN (select U1.SA1MYProgramID, U1.[MYProgramID Count],SameYear, CYEventDate
				,U1.[Year], U1.[Min Event Startdate],U1.[Max Event Startdate], ABS(DATEDIFF(DAY,U1.[Min Event Startdate],@CurrentDate)) AS MinD, ABS(DATEDIFF(DAY,U1.[Max Event Startdate],@CurrentDate)) AS MaxD
				,CASE WHEN SameYear = 'Different Year' then U1.CYEventDate
					  WHEN ABS(DATEDIFF(DAY,U1.[Min Event Startdate],@CurrentDate)) <= ABS(DATEDIFF(DAY,U1.[Max Event Startdate],@CurrentDate)) THEN U1.[Min Event Startdate]
					  WHEN ABS(DATEDIFF(DAY,U1.[Min Event Startdate],@CurrentDate)) > ABS(DATEDIFF(DAY,U1.[Max Event Startdate],@CurrentDate)) THEN U1.[Max Event Startdate]
				 ELSE @CurrentDate END AS SA1EligibleDate
				from [cteMYMultipleProgramSameYear] U1) SYDY 
			ON SA.MYProgramID = SYDY.SA1MYProgramID AND SA.EventStart=SYDY.[SA1EligibleDate]
WHERE	SA.Year >= 2016 AND SA.TestProg IS NULL 
	AND FC.FirstApprovalDate IS NOT NULL 
	AND SA.Status NOT IN ('TURNDOWN','SUBMITTED FOR ANALYSIS','NEGOTIATION','COPIED - UPDATE REQUIRED')
	AND LC.FormTitle <> 'Legal Amendment'
	AND SA.Multi = 'YES'
	),

[cteDuplicateRenewalIDSameYear] AS ( 
SELECT [Impact Year],
				ReportingYear, 
				RenewalID, 
				Count(RenewalID) AS 'RowCount',
				(SELECT MAX(CAST(T2.[Event Start Date] AS DATE))
						from cteFinalScript T2
						WHERE (T2.RenewalID = A.RenewalID) AND (T2.ReportingYear = Year(T2.[Event Start Date]))) AS 'NewestStartDate'
			  FROM cteFinalScript A
  WHERE RenewalID IS NOT NULL AND  [Impact Year] = ReportingYear
  GROUP BY ReportingYear, RenewalID,[Impact Year]
  HAVING COUNT(RenewalID) > 1),

[cteRenewalID] AS ( 
SELECT	A.ReportingYear, 
					A.RenewalID, 
					Count(A.RenewalID) AS  'RowCount',
					SUBSTRING((SELECT (','+Rtrim(T1.[Impact Year])) 
							from cteFinalScript T1
							WHERE T1.RenewalID = A.RenewalID
							FOR XML PATH('')),2,1000) AS 'Impact Years',
					CAST(Max(A.[Event Start Date]) AS DATE) AS 'NewestStartDate'
  FROM cteFinalScript A
  LEFT JOIN cteDuplicateRenewalIDSameYear B ON A.RenewalID = B.RenewalID
  WHERE A.RenewalID IS NOT NULL AND B.RenewalID IS NULL
  GROUP BY A.ReportingYear, A.RenewalID),


[cteFinalEligibleRenewalID] AS (SELECT	A.ReportingYear,
					A.RenewalID, 
					A.[Impact Years], 
					A.[RowCount],
					A.NewestStartDate, 
					CASE WHEN [Impact Years] LIKE CONCAT('%',ReportingYear,'%') 
										THEN (  SUBSTRING((SELECT (','+Rtrim(T1.[SADOCID])) 
															from cteFinalScript T1
															WHERE T1.RenewalID = A.RenewalID AND T1.[Impact Year]=A.[ReportingYear]
															FOR XML PATH('')),2,1000))
								WHEN [Impact Years] NOT LIKE CONCAT('%',ReportingYear,'%') 
										THEN (  SUBSTRING((SELECT (','+Rtrim(T1.[SADOCID])) 
															from cteFinalScript T1
															WHERE T1.RenewalID = A.RenewalID AND T1.[Event Start Date]=A.[NewestStartDate]
															FOR XML PATH('')),2,1000))
								END AS 'EligibleSADOCID'
FROM cteRenewalID A
UNION
Select		A.ReportingYear,
					A.RenewalID,
					CAST(A.[Impact Year] AS varchar) AS 'Impact Year',
					A.[RowCount],
					A.NewestStartDate, 
					SUBSTRING((SELECT (','+Rtrim(T1.[SADOCID])) 
															from cteFinalScript T1
															WHERE T1.RenewalID = A.RenewalID AND T1.[Event Start Date]=A.[NewestStartDate]
								FOR XML PATH('')),2,1000)  AS 'EligibleSADOCID'
FROM cteDuplicateRenewalIDSameYear A)

--SELECT A.*, 
--CASE WHEN B.EligibleSADOCID IS NOT NULL THEN 'Keep'
--ELSE 'Remove' END AS 'EligibleRenewalID'
-- FROM cteFinalScript A
-- LEFT JOIN [cteFinalEligibleRenewalID] B ON B.EligibleSADOCID = A.SADOCID
-- ORDER BY RenewalID

Select A.* FROM cteFinalScript A
LEFT JOIN cteFinalEligibleRenewalID B ON A.SADOCID = B.EligibleSADOCID