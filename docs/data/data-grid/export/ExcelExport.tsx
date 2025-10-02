import { DataGridPremium, GridColDef, GridRowsProp } from '@mui/x-data-grid-premium';

const dataGenerator = (rowCount: number) => {
  const rows: any[] = [];

  // Generate random data for all fields
  for (let i = 0; i < rowCount; i += 1) {
    const newRow = {
      jobTitle: `Job Title ${i}`,
      recruitmentDate: new Date(2020, 8, 12),
      contract: 'full time',
      id: i,
      test1: 'full time',
      test2: 'full time',
      test3: 'full time',
      test4: 'full time',
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
    };
    rows.push(newRow);
  }
  return rows;
};

const rows: GridRowsProp = dataGenerator(50_000);

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
    field: 'test1',
    headerName: 'Test1',
    type: 'singleSelect',
    valueOptions: ['full time', 'part time', 'intern'],
    width: 150,
  },
  {
    field: 'test2',
    headerName: 'Test2',
    type: 'singleSelect',
    valueOptions: ['full time', 'part time', 'intern'],
    width: 150,
  },
  {
    field: 'test3',
    headerName: 'Test3',
    type: 'singleSelect',
    valueOptions: ['full time', 'part time', 'intern'],
    width: 150,
  },
  {
    field: 'test4',
    headerName: 'Test4',
    type: 'singleSelect',
    valueOptions: ['full time', 'part time', 'intern'],
    width: 150,
  },
  {
    field: 'childData.subJobTitle',
    headerName: 'Sub Job Title',
    width: 100,
    isExportChildColumn: true,
    valueGetter: (value) => value,
    valueFormatter: (value) => value,
  },
  {
    field: 'childData.subRecruitmentDate',
    headerName: 'Sub Recruitment Date',
    width: 150,
    isExportChildColumn: true,
    valueGetter: (value) => value,
    valueFormatter: (value) => value,
  },
  {
    field: 'childData.someArrayData',
    headerName: 'Some array data',
    width: 150,
    isExportChildColumn: true,
    valueGetter: (value) => value,
    valueFormatter: (value) => (value as string[])?.join(', '),
  },
];

export default function ExcelExport() {
  return (
    <div style={{ height: 300, width: '100%' }}>
      <DataGridPremium rows={rows} columns={columns} showToolbar />
    </div>
  );
}
