import React from 'react';
import type { ResolvedICMemo } from '../types/memo';

export interface CarriedStageInsertSection {
  sectionNumber: number;
  heading: string;
  text: string;
}

interface CarriedICMemoBodyProps {
  memo: ResolvedICMemo;
  insertedAfterSec15?: CarriedStageInsertSection[];
  particularsSectionNumber?: number;
  termsSectionNumber?: number;
}

const TokenVal: React.FC<{ ph: Record<string, string>; code: string; bold?: boolean }> = ({
  ph,
  code,
  bold = true,
}) => {
  const val = ph[code];
  if (!val || val.trim() === '') {
    return (
      <span className="bg-amber-100 text-amber-900 px-1 rounded font-mono text-[8.5pt] font-semibold">
        {code}
      </span>
    );
  }
  return <span className={bold ? 'font-bold' : ''}>{val}</span>;
};

export const CarriedICMemoBody: React.FC<CarriedICMemoBodyProps> = ({
  memo,
  insertedAfterSec15 = [],
  particularsSectionNumber = 16,
  termsSectionNumber = 17,
}) => {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;

  const defaultCibStatuses = [
    'STD',
    'SMA',
    'SS',
    'DF',
    'BL',
    'Willful Default (WD)',
    'Write-Off',
    'Non-Funded',
    'Stay Order',
    'TOTAL',
  ];

  const rawSec3Rows =
    bp?.section3Rows && bp.section3Rows.length > 0
      ? bp.section3Rows
      : [
          { sl: 'a.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'b.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'c.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        ];
  const hasSec3SubTotal = rawSec3Rows.some(
    (r) => /sub[-\s]*total/i.test(r.sl) || /sub[-\s]*total/i.test(r.nameOfConcern)
  );
  const hasSec3GrandTotal = rawSec3Rows.some(
    (r) => /grand\s*total/i.test(r.sl) || /grand\s*total/i.test(r.nameOfConcern)
  );

  const sec4Rows =
    bp?.section4Rows && bp.section4Rows.length > 0
      ? bp.section4Rows
      : defaultCibStatuses.map((s) => ({
          status: s,
          unitBorrower: '',
          unitGuarantor: '',
          otherBorrower: '',
          otherGuarantor: '',
          remarks: '',
        }));

  const sec5Rows =
    bp?.section5Rows && bp.section5Rows.length > 0
      ? bp.section5Rows
      : [{ sl: '', typeOfCase: '', caseNo: '', filingDate: '', amount: '', previousDate: '', nextDate: '', presentStatus: '' }];

  const defaultSec6 = [
    { sl: 'a.', authority: 'IC & CD', lastInspectionDate: '', objectionStatus: '', objections: '', branchReply: '' },
    { sl: 'b.', authority: 'DBI, Bangladesh Bank', lastInspectionDate: '', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'c.', authority: 'DOS, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'd.', authority: 'BFIU, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'e', authority: 'BRD, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
  ];
  const sec6Rows = bp?.section6Rows && bp.section6Rows.length > 0 ? bp.section6Rows : defaultSec6;

  const sec7Primary =
    bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0
      ? bp.section7PrimaryRows
      : [
          { particulars: 'i) Hypothecation of Stock as on ………………..', value: '' },
          { particulars: 'ii) Hypothecation of Machinery', value: '' },
          { particulars: 'iii) Assignment of Bills against Work Order', value: '' },
          { particulars: 'iv) Margin held against existing LC & BG liability as on……..', value: '' },
        ];

  const sec7Schedules =
    bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0
      ? bp.section7CollateralSchedules
      : [
          'Schedule-01: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-02: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-03: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-04: Full Schedule with owner name as per Mortgage Deed.',
        ];

  const sec7ValRows =
    bp?.section7ValuationRows && bp.section7ValuationRows.length > 0
      ? bp.section7ValuationRows
      : ['Schedule-01', 'Schedule-03', 'Schedule-02', 'Schedule-04', 'Building', 'Other', 'Grand Total'].map((s) => ({
          schedule: s,
          origBranchMv: '',
          origBranchFsv: '',
          origSurvMv: '',
          origSurvFsv: '',
          presBranchMv: '',
          presBranchFsv: '',
          presSurvMv: '',
          presSurvFsv: '',
          presAvMv: '',
          presAvFsv: '',
        }));

  const sec9Rows =
    bp?.section9Rows && bp.section9Rows.length > 0
      ? bp.section9Rows
      : ['a.', 'b.', 'c.', 'd.', 'e.', 'f.', 'g.'].map((sl) => ({
          sl,
          referenceNo: '',
          date: '',
          limit: '',
          remarks: '',
        }));

  const sp = bp?.section10Particulars;
  const dirRows =
    bp?.section10Directors && bp.section10Directors.length > 0
      ? bp.section10Directors
      : [{ name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' }];
  const bankDirRows =
    bp?.section10BankDirectors && bp.section10BankDirectors.length > 0
      ? bp.section10BankDirectors
      : [{ name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' }];
  const finRows =
    bp?.section11Rows && bp.section11Rows.length > 0
      ? bp.section11Rows
      : [
          {
            year: '',
            grossSales: bp?.section11FinancialsNote || 'Nil',
            grossProfit: '',
            netProfit: '',
            totalAsset: '',
            totalLiabilities: '',
            netWorth: '',
          },
          { year: '', grossSales: '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' },
          { year: '', grossSales: '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' },
        ];

  return (
    <div className="space-y-3">
      {/* Main Settlement Table */}
      <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
        <colgroup>
          <col className="w-[12%]" />
          <col className="w-[8%]" />
          <col className="w-[7%]" />
          <col className="w-[21%]" />
          <col className="w-[11%]" />
          <col className="w-[11%]" />
          <col className="w-[11%]" />
          <col className="w-[19%]" />
        </colgroup>
        <thead>
          <tr>
            <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">Mode</th>
            <th colSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">Limit</th>
            <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">Particulars</th>
            <th className="border border-dotted border-black px-1.5 py-0.5 text-center">
              O/S as on <TokenVal ph={ph} code="#68#" />
            </th>
            <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Proposed Waiver</th>
            <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Settlement Amount</th>
            <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">Remarks</th>
          </tr>
          <tr>
            <th className="border border-dotted border-black px-1 py-0.5 text-center">Original</th>
            <th className="border border-dotted border-black px-1 py-0.5 text-center">RSD</th>
            <th className="border border-dotted border-black px-1 py-0.5 text-center">01</th>
            <th className="border border-dotted border-black px-1 py-0.5 text-center">02</th>
            <th className="border border-dotted border-black px-1 py-0.5 text-center">03 (1-2)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#5#" />
            </td>
            <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#6#" />
            </td>
            <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center">
              {bp?.limitRSD || '_'}
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Principal</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#7#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 font-bold leading-snug">
              <div>1)Settlement Deadline:</div>
              <div className="mb-2"><TokenVal ph={ph} code="#28#" /></div>
              <div>2) Break up</div>
              <div>Principal: <TokenVal ph={ph} code="#7#" /></div>
              <div>Charges: <TokenVal ph={ph} code="#8#" /></div>
              <div>COF: <TokenVal ph={ph} code="#31#" /></div>
              <div className="mb-2">Total: <TokenVal ph={ph} code="#32#" /></div>
              <div>3) Security Value:</div>
              <div>MV: <TokenVal ph={ph} code="#33#" /></div>
              <div>FSV: <TokenVal ph={ph} code="#34#" /></div>
            </td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Charges (Legal &amp; Others)</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#8#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Gestation Rent</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Charged Profit</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#10#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Compensation</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#11#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Uncharged Compensation</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#12#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
          <tr>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Uncharged Profit (for T/L)</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
          </tr>
          <tr>
            <td colSpan={4} className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
              Grand-Total
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#25#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#4#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
              <TokenVal ph={ph} code="#27#" />
            </td>
            <td className="border border-dotted border-black px-1.5 py-0.5"></td>
          </tr>
        </tbody>
      </table>

      {/* Down Payment Note */}
      <div className="border border-dotted border-black px-2 py-0.5 text-[10pt] font-bold">
        Note: Information Regarding Down payment (deposited Tk.<TokenVal ph={ph} code="#35#" />{' '}
        Million in A/C: <TokenVal ph={ph} code="#36#" /> on <TokenVal ph={ph} code="#68#" />.)
      </div>

      {/* Key Points */}
      <div className="space-y-1">
        <p className="font-bold underline text-[10pt]">Key Points:</p>
        <p className="text-[10pt]">
          This Memo is placed in conformity with{' '}
          <span className="font-bold">BRPD 1 Circular Letter No. 23 dated 29.06.2026</span> under the following contexts:
        </p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[5%]" />
            <col className="w-[42%]" />
            <col className="w-[11%]" />
            <col className="w-[5%]" />
            <col className="w-[22%]" />
            <col className="w-[15%]" />
          </colgroup>
          <thead>
            <tr>
              <th colSpan={3} className="border border-dotted border-black px-2 py-0.5 text-left">Particulars</th>
              <th colSpan={3} className="border border-dotted border-black px-2 py-0.5 text-left">Waiver recommended by:</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5">a.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Total [Profit+ Compensation+ Uncharged Comp.&amp; Profit]</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#42#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5">d.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Client</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#45#" /></td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Cost of Fund</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#31#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Branch</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#45#" /></td>
            </tr>
            <tr>
              <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">c.</td>
              <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">Maximum Waiver Considering COF [a-b]</td>
              <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#44#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">SAMD, HO</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#47#" /></td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5">g.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">HOIC</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#47#" /></td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5">h.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">MEC</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#47#" /></td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5">i.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">IC &amp; CD</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">: <TokenVal ph={ph} code="#47#" /></td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Justification */}
      <div className="space-y-1 text-[10pt]">
        <p className="font-bold underline">Justification of The Proposed Facility:</p>
        <div className="pl-2 space-y-0.5">
          <p className="whitespace-pre-line">
            a. <TokenVal ph={ph} code="#51#" bold={false} />
          </p>
          <p>b. Investment was not created against any fraud and forgery.</p>
          <p>c. The client is not a willful defaulter.</p>
        </div>
      </div>

      <p className="font-bold text-[10pt]">
        Previous Board Decision(s) Regarding Rescheduling/Waiver/Written-Off: {bp?.previousBoardDecisions || 'Nil'}
      </p>

      {/* Section 1 */}
      <div className="border border-dotted border-black px-2 py-0.5 font-bold text-[10pt]">
        1. Nature of Business: <TokenVal ph={ph} code="#54#" />
      </div>

      {/* Section 2 */}
      <div className="space-y-1">
        <p className="font-bold text-[10pt]">
          2. Liability Position of the Clients and Its Sister Concerns with Our Bank as on{' '}
          <TokenVal ph={ph} code="#68#" />:
        </p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[14%]" />
            <col className="w-[17%]" />
            <col className="w-[10%]" />
            <col className="w-[11%]" />
            <col className="w-[11%]" />
            <col className="w-[10%]" />
            <col className="w-[17%]" />
            <col className="w-[10%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Concern</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Mode of Investment</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Limit</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Outstanding</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Validity</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Overdue</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Recovery in existing deals till date [Mode wise]</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Remarks</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                <TokenVal ph={ph} code="#2#" />
              </td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                <div><TokenVal ph={ph} code="#55#" /></div>
                <div>Funded</div>
              </td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#56#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#25#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#58#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#59#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#60#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5">{bp?.sec2Remarks || ''}</td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">Non Funded</td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              <td className="border border-dotted border-black px-1.5 py-0.5">{memo.sec2NonFundedNote || ''}</td>
            </tr>
            <tr>
              <td colSpan={2} className="border border-dotted border-black px-1.5 py-0.5 font-bold text-center">Total</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#56#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#25#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#58#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#59#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold"><TokenVal ph={ph} code="#60#" /></td>
              <td className="border border-dotted border-black px-1.5 py-0.5"></td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Section 3 */}
      <div className="space-y-1">
        <p className="font-bold underline text-[10pt]">
          3. Liability of the Client &amp; Its Sister/Allied Concerns with Other Units of SIB:
        </p>
        <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[5%]" />
            <col className="w-[13%]" />
            <col className="w-[17%]" />
            <col className="w-[11%]" />
            <col className="w-[8%]" />
            <col className="w-[17%]" />
            <col className="w-[11%]" />
            <col className="w-[10%]" />
            <col className="w-[8%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">SL</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of Concern</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Banks/NBFIs</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Nature of Facility</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Limit</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Outstanding with profit/Interest</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Overdue/EOL</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Expiry Date</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">CL Status</th>
            </tr>
          </thead>
          <tbody>
            {rawSec3Rows.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.nameOfConcern}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.bankOrNbfi}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.natureOfFacility}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.limit}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.outstanding}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.overdueEol}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.expiryDate}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.clStatus}</td>
              </tr>
            ))}
            {!hasSec3SubTotal && (
              <tr>
                <td colSpan={4} className="border border-dotted border-black px-1 py-0.5 font-bold text-center">Sub-Total</td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
              </tr>
            )}
            {!hasSec3GrandTotal && (
              <tr>
                <td colSpan={4} className="border border-dotted border-black px-1 py-0.5 font-bold text-center">Grand Total</td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
                <td className="border border-dotted border-black px-1 py-0.5"></td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Section 4 */}
      <div className="space-y-1">
        <p className="font-bold text-[10pt]">
          <span className="underline">4. CIB Status:</span> Report dated <TokenVal ph={ph} code="#62#" /> shows:
        </p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[18%]" />
            <col className="w-[13%]" />
            <col className="w-[13%]" />
            <col className="w-[13%]" />
            <col className="w-[13%]" />
            <col className="w-[30%]" />
          </colgroup>
          <thead>
            <tr>
              <th rowSpan={2} className="border border-dotted border-black px-1 py-0.5 text-left">STATUS</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">With respective unit</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">With Other Bank</th>
              <th rowSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Remarks</th>
            </tr>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
            </tr>
          </thead>
          <tbody>
            {sec4Rows.map((r, i) => (
              <tr key={i} className={/TOTAL/i.test(r.status) ? 'font-bold' : ''}>
                <td className="border border-dotted border-black px-1.5 py-0.5">{r.status}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitBorrower}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitGuarantor}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherBorrower}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherGuarantor}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Section 5 */}
      <div className="space-y-1">
        <p className="font-bold text-[10pt]">
          <span className="underline">5. Legal Action Status as on Memo Date:</span> {bp?.section5StatusNote || ''}
        </p>
        <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[5%]" />
            <col className="w-[16%]" />
            <col className="w-[13%]" />
            <col className="w-[12%]" />
            <col className="w-[12%]" />
            <col className="w-[12%]" />
            <col className="w-[12%]" />
            <col className="w-[18%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Sl</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Type of Case</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Case No.</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Filing date</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Amount</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Previous date</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Next date</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Present status</th>
            </tr>
          </thead>
          <tbody>
            {sec5Rows.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.typeOfCase}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.caseNo}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.filingDate}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.amount}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.previousDate}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.nextDate}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.presentStatus}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Section 6 */}
      <div className="space-y-1">
        <p className="font-bold text-[10pt]">
          <span className="underline">6. Audit, Compliance and Regulatory Authority Related Status:</span>{' '}
          {bp?.section6StatusNote || ''}
        </p>
        <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
          <colgroup>
            <col className="w-[5%]" />
            <col className="w-[19%]" />
            <col className="w-[16%]" />
            <col className="w-[14%]" />
            <col className="w-[23%]" />
            <col className="w-[23%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">SL</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Authority</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Last Inspection Date</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Objection Status</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Objections (If yes)</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Branch Reply / Position</th>
            </tr>
          </thead>
          <tbody>
            {sec6Rows.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.authority}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.lastInspectionDate}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.objectionStatus}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.objections}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{r.branchReply}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Section 7 */}
      <div className="space-y-1.5 text-[10pt]">
        <p className="font-bold underline">7. Security:</p>
        <p className="font-bold pl-2">
          <span className="underline">a. Primary:</span> {bp?.section7PrimaryNote || ''}
        </p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[75%]" />
            <col className="w-[25%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-2 py-0.5 text-left">Particulars</th>
              <th className="border border-dotted border-black px-2 py-0.5 text-center">Value</th>
            </tr>
          </thead>
          <tbody>
            {sec7Primary.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-2 py-0.5">{r.particulars}</td>
                <td className="border border-dotted border-black px-2 py-0.5 text-center">{r.value}</td>
              </tr>
            ))}
            <tr className="font-bold">
              <td className="border border-dotted border-black px-2 py-0.5 text-center">Total value of primary security</td>
              <td className="border border-dotted border-black px-2 py-0.5 text-center">{bp?.section7PrimaryTotal || ''}</td>
            </tr>
          </tbody>
        </table>

        <p className="font-bold pl-2">
          <span className="underline">b. Collateral:</span> {bp?.section7CollateralNote || ''}
        </p>
        <div className="pl-2 space-y-0.5 text-[10pt]">
          {sec7Schedules.map((s, i) => (
            <p key={i}>{s}</p>
          ))}
        </div>

        <table className="w-full table-fixed text-[9pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[13%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
            <col className="w-[8.7%]" />
          </colgroup>
          <thead>
            <tr>
              <th rowSpan={2} className="border border-dotted border-black px-1 py-1 text-center">Schedule</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Orig. Branch</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Orig. Surveyor</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Pres. Branch</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Pres. Surveyor</th>
              <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">Pres. AV Cell</th>
            </tr>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
            </tr>
          </thead>
          <tbody>
            {sec7ValRows.map((r, i) => (
              <tr key={i} className={/Grand/i.test(r.schedule) ? 'font-bold' : ''}>
                <td className="border border-dotted border-black px-1 py-0.5">{r.schedule}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origBranchMv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origBranchFsv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origSurvMv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origSurvFsv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presBranchMv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presBranchFsv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presSurvMv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presSurvFsv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presAvMv}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presAvFsv}</td>
              </tr>
            ))}
            <tr>
              <td colSpan={11} className="border border-dotted border-black px-2 py-0.5 text-center font-bold">
                Outstanding to Lowest Security value Coverage (%): {bp?.section7CoveragePct || '—'}
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Section 8 */}
      <div className="space-y-1 text-[10pt]">
        <p className="font-bold underline">8. Source of Repayment:</p>
        <p className="whitespace-pre-line">{memo.sourceOfRepayment || '—'}</p>
      </div>

      {/* Section 9 */}
      <div className="space-y-1">
        <p className="font-bold underline text-[10pt]">9. Sequence of Head Office Sanction(S):</p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[7%]" />
            <col className="w-[36%]" />
            <col className="w-[16%]" />
            <col className="w-[15%]" />
            <col className="w-[26%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1.5 py-0.5 text-center">SL</th>
              <th className="border border-dotted border-black px-1.5 py-0.5 text-left">HO Sanction Reference No.</th>
              <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Date</th>
              <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Limit</th>
              <th className="border border-dotted border-black px-1.5 py-0.5 text-left">Remarks</th>
            </tr>
          </thead>
          <tbody>
            {sec9Rows.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.sl}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5">{r.referenceNo}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.date}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.limit}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Section 10 */}
      <div className="space-y-2">
        <p className="font-bold underline text-[10pt]">10. Brief Particulars of The Client:</p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
          <colgroup>
            <col className="w-[6%]" />
            <col className="w-[37%]" />
            <col className="w-[4%]" />
            <col className="w-[53%]" />
          </colgroup>
          <tbody>
            {[
              ['a.', 'Name of the Client', sp?.nameOfClient || ''],
              ['b.', 'Unique Customer ID', sp?.uniqueCustomerId || ''],
              ['c.', 'AWCD/MSD Account with opening date', sp?.awcdMsdAccount || ''],
              ['d.', 'Business Type/Nature', sp?.businessTypeNature || ''],
              ['e.', 'Business Management Type', sp?.businessManagementType || ''],
              ['f.', 'Client Type', sp?.clientType || ''],
              ['g.', 'Key Person with Phone No.', sp?.keyPersonPhone || ''],
              ['g.', 'Address', sp?.address || ''],
              ['h.', 'Office/Business Address', sp?.officeBusinessAddress || ''],
              ['h.', 'Factory Address', sp?.factoryAddress || ''],
              ['i.', 'Purpose of Investment', sp?.purposeOfInvestment || ''],
              ['j.', 'Limit of Investment', sp?.limitOfInvestment || ''],
              ['k.', 'Mode of Investment', sp?.modeOfInvestment || ''],
              ['l.', 'Last Visit Date', sp?.lastVisitDate || ''],
              ['m.', 'Stock Position', sp?.stockPosition || ''],
              ['n.', 'Fixed Asset', sp?.fixedAsset || ''],
              ['o.', 'Business Net Worth', sp?.businessNetWorth || ''],
              ['p.', 'Insurance Coverage', sp?.insuranceCoverage || ''],
              ['q.', 'Trade License & Validity', sp?.tradeLicenseValidity || ''],
              ['r.', 'Other Regulatory Documents', sp?.otherRegulatoryDocuments || ''],
              ['s.', 'Date of Original Investment', sp?.dateOfOriginalInvestment || ''],
              ['t.', '1st Disbursement Manager', sp?.firstDisbursementManager || ''],
              ['u.', 'First Date of Classification with Status', sp?.firstDateOfClassification || ''],
            ].map(([idx, label, val], i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{idx}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5">{label}</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                <td className="border border-dotted border-black px-1.5 py-0.5">{val}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <p className="font-bold underline text-[10pt]">• Particular of the Proprietor/Partner/Directors:</p>
        <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[22%]" />
            <col className="w-[8%]" />
            <col className="w-[14%]" />
            <col className="w-[15%]" />
            <col className="w-[12%]" />
            <col className="w-[14%]" />
            <col className="w-[15%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Age</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Education</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Designation</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Share (%)</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Experience</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Net worth</th>
            </tr>
          </thead>
          <tbody>
            {dirRows.map((d, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.age}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{d.education}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.sharePct}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.experience}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.netWorth}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <p className="font-bold underline text-[10pt]">• Remarks: Director of any Bank/ FI/ Insurance company:</p>
        <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[22%]" />
            <col className="w-[22%]" />
            <col className="w-[18%]" />
            <col className="w-[19%]" />
            <col className="w-[19%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Director</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the bank/NBFIs</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-left">Designation/Position</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Percentage of Shares</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Number of Shares</th>
            </tr>
          </thead>
          <tbody>
            {bankDirRows.map((d, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{d.bankNbfi}</td>
                <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.sharePct}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.numShares}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Section 11 */}
      <div className="space-y-1">
        <p className="font-bold underline text-[10pt]">11. Financials of The Client:</p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black tabular-nums">
          <colgroup>
            <col className="w-[10%]" />
            <col className="w-[15%]" />
            <col className="w-[15%]" />
            <col className="w-[15%]" />
            <col className="w-[15%]" />
            <col className="w-[15%]" />
            <col className="w-[15%]" />
          </colgroup>
          <thead>
            <tr>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Year</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Gross Sales</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Gross Profit</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Net Profit</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Total Asset</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Total liabilities</th>
              <th className="border border-dotted border-black px-1 py-0.5 text-center">Net Worth</th>
            </tr>
          </thead>
          <tbody>
            {finRows.map((r, i) => (
              <tr key={i}>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.year}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.grossSales}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.grossProfit}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.netProfit}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.totalAsset}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.totalLiabilities}</td>
                <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.netWorth}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Sections 12–15 */}
      <div className="space-y-1.5 text-[10pt]">
        <div>
          <p className="font-bold underline">12. Contents of the Client’s Prayer:</p>
          <p className="whitespace-pre-line">{memo.clientPrayer || '—'}</p>
        </div>
        <div>
          <p className="font-bold underline">12. Branch Visit Report:</p>
          <p className="whitespace-pre-line">{memo.branchVisitReport || '—'}</p>
        </div>
        <div>
          <p>
            <span className="font-bold underline">13. Recommendation of The Branch:</span> The Branch Recommended the following vide Proposal Ref: <TokenVal ph={ph} code="#64#" /> dated <TokenVal ph={ph} code="#65#" />
          </p>
          {memo.sec13BranchRecommendation && (
            <p className="pl-4 whitespace-pre-line">{memo.sec13BranchRecommendation}</p>
          )}
        </div>
        <div>
          <p>
            <span className="font-bold underline">14. Recommendation of Internal Control &amp; Compliance Division:</span>{' '}
            <span className="whitespace-pre-line">{memo.sec14IccdRecommendation || 'N/A'}</span>
          </p>
        </div>
        <div>
          <p className="font-bold underline">15. Observation and Recommendation of SAMD:</p>
          <div className="pl-4 whitespace-pre-line">{memo.samdObservation || '—'}</div>
        </div>
        {insertedAfterSec15.map((sec) => (
          <div key={`${sec.sectionNumber}-${sec.heading}`}>
            <p className="font-bold underline">
              {sec.sectionNumber}. {sec.heading}:
            </p>
            <div className="pl-4 whitespace-pre-line">{sec.text || '—'}</div>
          </div>
        ))}
      </div>

      {/* Section 16 (or 17/18 in MEC/Board) */}
      <div className="space-y-1">
        <p className="font-bold underline text-[10pt]">
          {particularsSectionNumber}. Particulars of the facility:
        </p>
        <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
          <colgroup>
            <col className="w-[5%]" />
            <col className="w-[24%]" />
            <col className="w-[71%]" />
          </colgroup>
          <tbody>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">a.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Facility</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                EXIT with Waiver under BRPD 1 Circular Letter No. 23 dated 29.06.2026
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">b.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Outstanding</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                Tk. <TokenVal ph={ph} code="#25#" /> million as on <TokenVal ph={ph} code="#68#" />.
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">c.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Down Payment</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                Down-payment of Tk. <TokenVal ph={ph} code="#35#" /> Million is kept in{' '}
                <TokenVal ph={ph} code="#36#" /> A/C <TokenVal ph={ph} code="#36#" />
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">d.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Waiver</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                Tk. <TokenVal ph={ph} code="#4#" /> million as on <TokenVal ph={ph} code="#68#" />
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">e.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Settlement/Exit amount</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                Tk. <TokenVal ph={ph} code="#27#" /> Million as on <TokenVal ph={ph} code="#68#" />
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">f.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Settlement period</td>
              <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                <TokenVal ph={ph} code="#28#" />
              </td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">g.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Profit/Rent rate</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Without charging further profit.</td>
            </tr>
            <tr>
              <td className="border border-dotted border-black px-1.5 py-0.5 text-center">h.</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Repayment Method</td>
              <td className="border border-dotted border-black px-1.5 py-0.5">Lump sum.</td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Section 17 (or 18/19 in MEC/Board) */}
      <div className="space-y-1 text-[10pt]">
        <p className="font-bold underline">
          {termsSectionNumber}. Special Term(s) &amp; Condition(s):
        </p>
        <p className="pl-4">
          If the proposal approved, the branch is advised to comply the following special terms and conditions:
        </p>
        <ol className="pl-10 space-y-1 list-none leading-relaxed text-[9.5pt]">
          <li>i. The waiver of profit shall be adjusted from FAD, Head Office from the provision amount kept (if any) and the rest amount shall be adjusted by branch gradually after adjustment of the settlement amount within the approved time frame.</li>
          <li>ii. The final settlement amount will be determined at the date of execution of settlement considering up to date liability position.</li>
          <li>iii. The Branch must ensure adjustment of all other direct and indirect liabilities of the client before allowing the final redemption of the mortgaged collateral security(s).</li>
          <li>iv. The Branch must ensure adjustment of all non-funded liabilities (L/C, Bank Guarantee, Credit Card etc.) before allowing final redemption of the mortgaged collateral security(s). The Branch must also ensure that any encashable security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client) including collateral is properly mentioned in the proposal. If any encashable security is found not mentioned in the proposal and later encashed in favor of the client, the Manager will be held responsible.</li>
          <li>v. If the client fails to repay or fails to adjust the liability within the given time or fails to pay installments regularly, all facilities shall be cancelled and legal action shall be initiated by the Branch.</li>
          <li>vi. Branch should be utmost sincere to protect bank interest in the course of execution partial redemption.</li>
          <li>vii. Branch is advised to follow the guidelines of Bangladesh Bank BRPD circular No. 06/22, Circular Letter No.18/22, 46/22 and subsequent up-dates as well as all instruction circular(s) of SIBL.</li>
          <li className="font-bold">viii. The Profit rate of the stipulated investment shall not be changed or reduced to 0.00% until full and final adjustment of the Investment, as approved by Head Office.</li>
          <li>ix. All the legal expenses and other charges (if any) already incurred and to be incurred in connection with the withdrawal of case/suits and others issues shall be borne by the client.</li>
        </ol>
      </div>
    </div>
  );
};
