import { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import Icon from '@/components/ui/icon';
import { useToast } from '@/hooks/use-toast';
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from 'recharts';

interface CellData {
  value: string | number;
  formula?: string;
}

interface SheetData {
  [key: string]: CellData[][];
}

const COLORS = ['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899'];

export default function ExcelCalculator() {
  const [sheetData, setSheetData] = useState<SheetData>({});
  const [activeSheet, setActiveSheet] = useState<string>('');
  const [isDragging, setIsDragging] = useState(false);
  const [fileName, setFileName] = useState<string>('');
  const [chartData, setChartData] = useState<any[]>([]);
  const { toast } = useToast();

  const processFile = useCallback((file: File) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        
        const sheets: SheetData = {};
        let firstSheet = '';
        
        workbook.SheetNames.forEach((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          
          sheets[sheetName] = (jsonData as any[]).map(row => 
            (row as any[]).map(cell => ({ value: cell }))
          );
          
          if (index === 0) {
            firstSheet = sheetName;
            generateChartData(jsonData as any[][]);
          }
        });
        
        setSheetData(sheets);
        setActiveSheet(firstSheet);
        setFileName(file.name);
        
        toast({
          title: "Файл загружен",
          description: `${file.name} успешно импортирован`,
        });
      } catch (error) {
        toast({
          title: "Ошибка",
          description: "Не удалось загрузить файл",
          variant: "destructive",
        });
      }
    };
    
    reader.readAsBinaryString(file);
  }, [toast]);

  const generateChartData = (data: any[][]) => {
    if (data.length < 2) return;
    
    const headers = data[0];
    const chartPoints = data.slice(1, 7).map((row, idx) => {
      const point: any = { name: row[0] || `Строка ${idx + 1}` };
      
      row.slice(1).forEach((value, colIdx) => {
        if (typeof value === 'number') {
          point[headers[colIdx + 1] || `Значение ${colIdx + 1}`] = value;
        }
      });
      
      return point;
    });
    
    setChartData(chartPoints.filter(p => Object.keys(p).length > 1));
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      processFile(file);
    } else {
      toast({
        title: "Неверный формат",
        description: "Пожалуйста, загрузите файл Excel (.xlsx или .xls)",
        variant: "destructive",
      });
    }
  }, [processFile, toast]);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const exportToExcel = () => {
    if (!activeSheet || !sheetData[activeSheet]) return;
    
    const ws = XLSX.utils.aoa_to_sheet(
      sheetData[activeSheet].map(row => row.map(cell => cell.value))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeSheet);
    XLSX.writeFile(wb, `${fileName || 'export'}_edited.xlsx`);
    
    toast({
      title: "Экспорт завершён",
      description: "Файл успешно сохранён",
    });
  };

  const currentData = activeSheet ? sheetData[activeSheet] : [];

  return (
    <div className="min-h-screen bg-background p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex items-center justify-between">
          <div>
            <h1 className="text-3xl font-bold text-foreground">Excel Калькулятор</h1>
            {fileName && (
              <p className="text-sm text-muted-foreground mt-1">{fileName}</p>
            )}
          </div>
          
          {currentData.length > 0 && (
            <Button onClick={exportToExcel} className="gap-2">
              <Icon name="Download" size={18} />
              Экспортировать
            </Button>
          )}
        </header>

        {currentData.length === 0 ? (
          <Card
            className={`border-2 border-dashed transition-colors ${
              isDragging ? 'border-primary bg-primary/5' : 'border-border'
            }`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
          >
            <div className="flex flex-col items-center justify-center py-20 px-6 text-center">
              <div className="w-20 h-20 rounded-full bg-primary/10 flex items-center justify-center mb-6">
                <Icon name="FileSpreadsheet" size={40} className="text-primary" />
              </div>
              
              <h2 className="text-2xl font-semibold mb-2">Загрузите Excel файл</h2>
              <p className="text-muted-foreground mb-6 max-w-md">
                Перетащите .xlsx или .xls файл сюда или нажмите кнопку для выбора
              </p>
              
              <label htmlFor="file-upload">
                <Button className="gap-2" asChild>
                  <span>
                    <Icon name="Upload" size={18} />
                    Выбрать файл
                  </span>
                </Button>
              </label>
              <input
                id="file-upload"
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileInput}
                className="hidden"
              />
            </div>
          </Card>
        ) : (
          <>
            {Object.keys(sheetData).length > 1 && (
              <div className="flex gap-2 flex-wrap">
                {Object.keys(sheetData).map((sheetName) => (
                  <Button
                    key={sheetName}
                    variant={activeSheet === sheetName ? 'default' : 'outline'}
                    onClick={() => {
                      setActiveSheet(sheetName);
                      const data = sheetData[sheetName];
                      generateChartData(data.map(row => row.map(cell => cell.value)));
                    }}
                    className="gap-2"
                  >
                    <Icon name="Sheet" size={16} />
                    {sheetName}
                  </Button>
                ))}
              </div>
            )}

            <Card className="overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full border-collapse">
                  <thead>
                    <tr className="bg-muted">
                      <th className="border border-border p-2 text-xs font-semibold text-center w-12 sticky left-0 bg-muted z-10">
                        #
                      </th>
                      {currentData[0]?.map((_, colIndex) => (
                        <th
                          key={colIndex}
                          className="border border-border p-2 text-xs font-semibold text-center min-w-[120px]"
                        >
                          {String.fromCharCode(65 + colIndex)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentData.map((row, rowIndex) => (
                      <tr key={rowIndex} className="hover:bg-muted/50 transition-colors">
                        <td className="border border-border p-2 text-xs font-semibold text-center bg-muted sticky left-0 z-10">
                          {rowIndex + 1}
                        </td>
                        {row.map((cell, cellIndex) => (
                          <td
                            key={cellIndex}
                            className="border border-border p-2 text-sm"
                          >
                            <div className="min-h-[20px]">
                              {typeof cell.value === 'number' 
                                ? cell.value.toLocaleString('ru-RU')
                                : cell.value}
                            </div>
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>

            {chartData.length > 0 && (
              <div className="grid md:grid-cols-2 gap-6">
                <Card className="p-6">
                  <h3 className="text-lg font-semibold mb-4 flex items-center gap-2">
                    <Icon name="BarChart3" size={20} />
                    Столбчатая диаграмма
                  </h3>
                  <ResponsiveContainer width="100%" height={300}>
                    <BarChart data={chartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                      <XAxis dataKey="name" tick={{ fontSize: 12 }} />
                      <YAxis tick={{ fontSize: 12 }} />
                      <Tooltip />
                      <Legend />
                      {Object.keys(chartData[0] || {})
                        .filter(key => key !== 'name')
                        .map((key, idx) => (
                          <Bar key={key} dataKey={key} fill={COLORS[idx % COLORS.length]} />
                        ))}
                    </BarChart>
                  </ResponsiveContainer>
                </Card>

                <Card className="p-6">
                  <h3 className="text-lg font-semibold mb-4 flex items-center gap-2">
                    <Icon name="LineChart" size={20} />
                    Линейный график
                  </h3>
                  <ResponsiveContainer width="100%" height={300}>
                    <LineChart data={chartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                      <XAxis dataKey="name" tick={{ fontSize: 12 }} />
                      <YAxis tick={{ fontSize: 12 }} />
                      <Tooltip />
                      <Legend />
                      {Object.keys(chartData[0] || {})
                        .filter(key => key !== 'name')
                        .map((key, idx) => (
                          <Line
                            key={key}
                            type="monotone"
                            dataKey={key}
                            stroke={COLORS[idx % COLORS.length]}
                            strokeWidth={2}
                          />
                        ))}
                    </LineChart>
                  </ResponsiveContainer>
                </Card>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
