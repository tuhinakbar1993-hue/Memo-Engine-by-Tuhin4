import type { MemoEngineDefinition } from '../types/memo';

export const MEMO_ENGINES: MemoEngineDefinition[] = [
  {
    id: 'ic',
    name: 'IC Memo',
    status: 'implemented',
    description: 'Investment Committee Memorandum for Exit with Waiver under BRPD Circular Letter No. 23.',
  },
  {
    id: 'mec',
    name: 'MEC Memo',
    status: 'implemented',
    description: 'Management Executive Committee Memorandum carrying forward full IC Memo data with Chronologic Order of the IC Memo, Date & Recommendation.',
  },
  {
    id: 'board',
    name: 'Board Memo',
    status: 'implemented',
    description: 'Board of Directors Memorandum carrying forward IC & MEC data with Chronologic Order of the MEC Memo, Date & Recommendation.',
  },
  {
    id: 'sanction',
    name: 'Sanction',
    status: 'implemented',
    description: 'Head Office Sanction Letter carrying forward Chronologic Order of the Board Memo, Date, cross-verified figures & operative terms.',
  },
];
