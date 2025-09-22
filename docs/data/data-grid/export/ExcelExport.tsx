import { DataGridPremium, GridColDef, GridRowsProp } from '@mui/x-data-grid-premium';

const rows: GridRowsProp = [
  {
    jobTitle: 'Head of Human Resources',
    recruitmentDate: new Date(2020, 8, 12),
    contract: 'full time',
    id: 0,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Head of Sales',
    recruitmentDate: new Date(2017, 3, 4),
    contract: 'full time',
    id: 1,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2020, 11, 20),
    contract: 'full time',
    id: 2,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2020, 10, 14),
    contract: 'part time',
    id: 3,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2017, 10, 29),
    contract: 'part time',
    id: 4,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2020, 7, 21),
    contract: 'full time',
    id: 5,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2020, 7, 20),
    contract: 'intern',
    id: 6,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Sales Person',
    recruitmentDate: new Date(2019, 6, 28),
    contract: 'full time',
    id: 7,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Head of Engineering',
    recruitmentDate: new Date(2016, 3, 14),
    contract: 'full time',
    id: 8,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Tech lead front',
    recruitmentDate: new Date(2016, 5, 17),
    contract: 'full time',
    id: 9,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Front-end developer',
    recruitmentDate: new Date(2019, 11, 7),
    contract: 'full time',
    id: 10,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Tech lead devops',
    recruitmentDate: new Date(2021, 7, 1),
    contract: 'full time',
    id: 11,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Tech lead back',
    recruitmentDate: new Date(2017, 0, 12),
    contract: 'full time',
    id: 12,
    childData: [
      {
        jobTitle: 'Back-end developer',
        recruitmentDate: new Date(2020, 1, 3),
        contract: 'full time',
        id: '12-1',
      },
      {
        jobTitle: 'Back-end developer',
        recruitmentDate: new Date(2019, 3, 23),
        contract: 'full time',
        id: '12-2',
      },
    ],
  },
  {
    jobTitle: 'Back-end developer',
    recruitmentDate: new Date(2019, 2, 22),
    contract: 'intern',
    id: 13,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
  {
    jobTitle: 'Back-end developer',
    recruitmentDate: new Date(2018, 4, 19),
    contract: 'part time',
    id: 14,
    childData: [
      {
        subJobTitle: 'Back-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
      {
        subJobTitle: 'Front-end developer',
        subRecruitmentDate: new Date(2020, 1, 3).toLocaleDateString(),
        someArrayData: ['One', 'Two', 'Three'],
      },
    ],
  },
];

const columns: GridColDef[] = [
  { field: 'jobTitle', headerName: 'Job Title', width: 200 },
  {
    field: 'recruitmentDate',
    headerName: 'Recruitment Date',
    type: 'date',
    width: 150,
  },
  {
    field: 'contract',
    headerName: 'Contract Type',
    type: 'singleSelect',
    valueOptions: ['full time', 'part time', 'intern'],
    width: 150,
  },
  {
    field: 'childData.subJobTitle',
    headerName: 'Sub Job Title',
    width: 100,
    isExportChildColumn: true,
    valueGetter: (value, row) => value,
    valueFormatter: (value, row) => value,
  },
  {
    field: 'childData.subRecruitmentDate',
    headerName: 'Sub Recruitment Date',
    width: 150,
    isExportChildColumn: true,
    valueGetter: (value, row) => value,
    valueFormatter: (value, row) => value,
  },
  {
    field: 'childData.someArrayData',
    headerName: 'Some array data',
    width: 150,
    isExportChildColumn: true,
    valueGetter: (value, row) => value,
    valueFormatter: (value, row) => (value as string[])?.join(', '),
  },
];

export default function ExcelExport() {
  return (
    <div style={{ height: 300, width: '100%' }}>
      <DataGridPremium rows={rows} columns={columns} showToolbar />
    </div>
  );
}
