import React, { useState, useCallback, useMemo } from "react";
import { createRoot } from "react-dom/client";

// TypeScript declaration for the XLSX library loaded from CDN
declare var XLSX: any;

// Utility to format dates into DD/MM/YYYY
const formatDate = (date: Date | any): string => {
  if (date instanceof Date && !isNaN(date.getTime())) {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }
  if (typeof date === "string") return date;
  return "";
};

// Column headers for the final output (moved outside component for stability)
const KHMER_HEADERS = {
  id: "លេខរៀង",
  title: "ចំណងជើង",
  pageView: "អ្នកទស្សនា",
  category: "ប្រភេទ",
  poster: "អ្នកបញ្ចូល",
  reporter: "អ្នករាយការណ៍",
  location: "អ្នករាយការណ៍តំបន់",
  date: "កាលបរិច្ឆេទ",
};

const VIEW_FILTERS = [
  { key: "all", label: "ទាំងអស់" },
  { key: "under100", label: "ក្រោម 100" },
  { key: "under500", label: "ក្រោម 500" },
  { key: "between500and1000", label: "ចន្លោះ 500-1000" },
  { key: "over1000", label: "លើស 1000" },
];

const App = () => {
  // State for file data and app status
  const [noPosterFile, setNoPosterFile] = useState<{
    name: string;
    data: any[];
  } | null>(null);
  const [posterFile, setPosterFile] = useState<{
    name: string;
    data: any[];
  } | null>(null);
  const [mergedData, setMergedData] = useState<any[] | null>(null);
  const [dateRange, setDateRange] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState<string | null>(null);

  // State for filtering
  const [viewFilter, setViewFilter] = useState("all");

  const HEADER_MAP_NO_POSTER: { [key: string]: string } = {
    ID: "id",
    Title: "title",
    PageView: "pageView",
    Category: "category",
    Reporter: "reporter",
    Location: "location",
    "Date Publi": "date",
  };

  const HEADER_MAP_POSTER: { [key: string]: string } = {
    ID: "id",
    Title: "title",
    Category: "category",
    Poster: "poster",
    Reporter: "reporter",
    Location: "location",
    Date: "date",
  };

  const handleFileUpload = useCallback(
    (file: File, type: "no-poster" | "poster") => {
      if (!file) return;

      setIsLoading(true);
      setError(null);
      setMergedData(null);
      setDateRange(null);
      setViewFilter("all");

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: "binary", cellDates: true });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

          if (type === "no-poster") {
            setNoPosterFile({ name: file.name, data: jsonData });
          } else {
            setPosterFile({ name: file.name, data: jsonData });
          }
        } catch (err) {
          console.error(err);
          setError(
            `មានបញ្ហាក្នុងការអានឯកសារ ${file.name}។ សូមប្រាកដថាវាជាឯកសារ Excel ដែលត្រឹមត្រូវ។`
          );
        } finally {
          setIsLoading(false);
        }
      };
      reader.onerror = () => {
        setError(`មានបញ្ហាក្នុងការអានឯកសារ ${file.name}។`);
        setIsLoading(false);
      };
      reader.readAsBinaryString(file);
    },
    []
  );

  const processFiles = useCallback(() => {
    if (!noPosterFile || !posterFile) {
      setError("សូមផ្ទុកឡើងឯកសារទាំងពីរមុននឹងដំណើរការ។");
      return;
    }

    setIsLoading(true);
    setError(null);

    setTimeout(() => {
      try {
        const posterDataMap = new Map();
        posterFile.data.forEach((row: any) => {
          const newRow: any = {};
          for (const key in HEADER_MAP_POSTER) {
            if (row[key] !== undefined) {
              newRow[HEADER_MAP_POSTER[key]] = row[key];
            }
          }
          newRow.pageView = row["PageView"] || 0;
          const dateValue = newRow.date;
          if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
            newRow._dateObj = dateValue;
            newRow.date = formatDate(dateValue);
          } else {
            newRow._dateObj = null;
            newRow.date = dateValue ? String(dateValue) : "";
          }
          if (newRow.id) {
            posterDataMap.set(String(newRow.id), newRow);
          }
        });

        const noPosterIdSet = new Set<string>();
        const processedData = noPosterFile.data.map((row: any) => {
          const newRow: any = {};
          for (const key in HEADER_MAP_NO_POSTER) {
            if (row[key] !== undefined) {
              newRow[HEADER_MAP_NO_POSTER[key]] = row[key];
            }
          }

          const dateValue = newRow.date;
          if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
            newRow._dateObj = dateValue;
            newRow.date = formatDate(dateValue);
          } else {
            newRow._dateObj = null;
            newRow.date = dateValue ? String(dateValue) : "";
          }

          if (newRow.id) {
            const stringId = String(newRow.id);
            noPosterIdSet.add(stringId);

            const posterInfo = posterDataMap.get(stringId);
            if (posterInfo) {
              newRow.poster = posterInfo.poster;
              newRow.reporter = posterInfo.reporter;
              newRow.location = posterInfo.location;
              newRow.category = posterInfo.category;
              newRow.date = posterInfo.date;
              newRow._dateObj = posterInfo._dateObj;
            } else {
              newRow.poster = "ស្រីណា";
            }
          } else {
            newRow.poster = "ស្រីណា";
          }
          return newRow;
        });

        posterDataMap.forEach((posterRow, id) => {
          if (!noPosterIdSet.has(id)) {
            processedData.push(posterRow);
          }
        });

        processedData.sort((a, b) => (b.pageView || 0) - (a.pageView || 0));

        const validDates = processedData
          .map((row) => row._dateObj)
          .filter((d): d is Date => d instanceof Date && !isNaN(d.getTime()));

        if (validDates.length > 0) {
          const minDate = new Date(
            Math.min(...validDates.map((d) => d.getTime()))
          );
          const maxDate = new Date(
            Math.max(...validDates.map((d) => d.getTime()))
          );
          setDateRange(`${formatDate(minDate)} - ${formatDate(maxDate)}`);
        } else {
          setDateRange(null);
        }

        setMergedData(processedData);
      } catch (err) {
        console.error(err);
        setError(
          "មានបញ្ហាកើតឡើងកំឡុងពេលដំណើរការទិន្នន័យ។ សូមពិនិត្យមើលថាតើជួរឈររបស់ឯកសារត្រូវគ្នានឹងទម្រង់ដែលរំពឹងទុកដែរឬទេ។"
        );
      } finally {
        setIsLoading(false);
      }
    }, 100);
  }, [noPosterFile, posterFile]);

  const filteredData = useMemo(() => {
    if (!mergedData) return null;

    let data = [...mergedData];

    switch (viewFilter) {
      case "under100":
        data = data.filter((row) => (row.pageView || 0) < 100);
        break;
      case "under500":
        data = data.filter((row) => (row.pageView || 0) < 500);
        break;
      case "between500and1000":
        data = data.filter(
          (row) => (row.pageView || 0) >= 500 && (row.pageView || 0) <= 1000
        );
        break;
      case "over1000":
        data = data.filter((row) => (row.pageView || 0) > 1000);
        break;
      default:
        break;
    }

    return data;
  }, [mergedData, viewFilter]);

  const outputHeaders = useMemo(() => Object.values(KHMER_HEADERS), []);
  const outputKeys = useMemo(() => Object.keys(KHMER_HEADERS), []);

  const downloadExcel = useCallback(() => {
    if (!filteredData) return;

    const dataToExport = filteredData.map((row) => {
      const exportRow: any = {};
      outputKeys.forEach((key) => {
        const header = KHMER_HEADERS[key as keyof typeof KHMER_HEADERS];
        if (key === "date") {
          // Prioritize the JS Date object for correct Excel cell type.
          if (
            row["_dateObj"] instanceof Date &&
            !isNaN(row["_dateObj"].getTime())
          ) {
            exportRow[header] = row["_dateObj"];
          } else {
            // Fallback to the string value if the Date object is not available.
            // This prevents empty cells.
            exportRow[header] = row[key] || null;
          }
        } else {
          exportRow[header] = row[key];
        }
      });
      return exportRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport, {
      cellDates: true,
    });

    // Define styles
    const headerStyle = { font: { name: "Kh Battambang", sz: 12, bold: true } };
    const defaultCellStyle = { font: { name: "Kh Battambang", sz: 12 } };
    const dateCellStyle = {
      font: { name: "Kh Battambang", sz: 12 },
      numFmt: "dd/mm/yyyy",
    };

    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const dateColumnIndex = outputKeys.indexOf("date");

    // Apply styles to all cells
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
        const cell = worksheet[cellRef];
        if (!cell) continue;

        if (R === 0) {
          // Header row
          cell.s = headerStyle;
        } else {
          // Data rows
          // Apply date-specific style or default style
          if (C === dateColumnIndex && cell.t === "d") {
            cell.s = dateCellStyle;
          } else {
            cell.s = defaultCellStyle;
          }
        }
      }
    }

    worksheet["!cols"] = [
      { wch: 10 },
      { wch: 50 },
      { wch: 12 },
      { wch: 20 },
      { wch: 15 },
      { wch: 20 },
      { wch: 20 },
      { wch: 15 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Data");
    XLSX.writeFile(workbook, "Filtered_Report.xlsx");
  }, [filteredData, outputKeys]);

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>, type: string) => {
    e.preventDefault();
    setDragOver(type);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(null);
  };

  const handleDrop = (
    e: React.DragEvent<HTMLDivElement>,
    type: "no-poster" | "poster"
  ) => {
    e.preventDefault();
    setDragOver(null);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      handleFileUpload(files[0], type);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col items-center p-4 sm:p-6 md:p-8">
      <main className="w-full bg-white rounded-xl shadow-lg p-6 sm:p-8">
        <header className="text-center mb-8">
          <h1 className="text-5xl font-bold text-gray-800">
            កម្មវិធីបម្លែងទិន្នន័យ Excel
          </h1>
        </header>

        {error && (
          <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded-lg text-center mb-6 text-xl">
            {error}
          </div>
        )}

        {!mergedData && (
          <>
            <section className="grid grid-cols-1 md:grid-cols-2 gap-10 mb-8">
              <div
                className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors duration-300 ${
                  dragOver === "no-poster"
                    ? "border-blue-500 bg-blue-50"
                    : "border-gray-300"
                }`}
                onDragOver={(e) => handleDragOver(e, "no-poster")}
                onDragLeave={handleDragLeave}
                onDrop={(e) => handleDrop(e, "no-poster")}
              >
                <h3 className="text-3xl font-semibold text-gray-700">
                  ឯកសារដែលគ្មាន Poster
                </h3>
                <p className="text-gray-500 my-4 text-xl">
                  ផ្ទុកឡើងឯកសារដែលគ្មានជួរឈរ Poster ។
                </p>
                <label
                  htmlFor="no-poster-file"
                  className="cursor-pointer text-blue-600 font-bold inline-block py-3 px-6 border border-blue-600 rounded-md hover:bg-blue-600 hover:text-white transition-colors text-xl"
                >
                  ជ្រើសរើសឯកសារ
                </label>
                <input
                  id="no-poster-file"
                  type="file"
                  className="hidden"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) =>
                    handleFileUpload(e.target.files?.[0] as File, "no-poster")
                  }
                />
                <div className="mt-4 text-lg text-gray-600 h-7">
                  {noPosterFile
                    ? `បានផ្ទុកឡើង៖ ${noPosterFile.name}`
                    : "មិនទាន់មានឯកសារ។"}
                </div>
              </div>
              <div
                className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors duration-300 ${
                  dragOver === "poster"
                    ? "border-blue-500 bg-blue-50"
                    : "border-gray-300"
                }`}
                onDragOver={(e) => handleDragOver(e, "poster")}
                onDragLeave={handleDragLeave}
                onDrop={(e) => handleDrop(e, "poster")}
              >
                <h3 className="text-3xl font-semibold text-gray-700">
                  ឯកសារដែលមាន Poster
                </h3>
                <p className="text-gray-500 my-4 text-xl">
                  ផ្ទុកឡើងឯកសារដែលមានជួរឈរ Poster ។
                </p>
                <label
                  htmlFor="poster-file"
                  className="cursor-pointer text-blue-600 font-bold inline-block py-3 px-6 border border-blue-600 rounded-md hover:bg-blue-600 hover:text-white transition-colors text-xl"
                >
                  ជ្រើសរើសឯកសារ
                </label>
                <input
                  id="poster-file"
                  type="file"
                  className="hidden"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) =>
                    handleFileUpload(e.target.files?.[0] as File, "poster")
                  }
                />
                <div className="mt-4 text-lg text-gray-600 h-7">
                  {posterFile
                    ? `បានផ្ទុកឡើង៖ ${posterFile.name}`
                    : "មិនទាន់មានឯកសារ។"}
                </div>
              </div>
            </section>

            <section className="text-center mb-8">
              <button
                onClick={processFiles}
                disabled={!noPosterFile || !posterFile || isLoading}
                className="bg-blue-600 text-white font-bold py-4 px-10 rounded-lg text-xl hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-all shadow-md hover:shadow-lg transform hover:-translate-y-0.5"
              >
                {isLoading ? "កំពុងដំណើរការ..." : "ដំណើរការឯកសារ"}
              </button>
            </section>
          </>
        )}

        {isLoading && (
          <div className="w-16 h-16 border-4 border-gray-200 border-t-blue-600 rounded-full animate-spin mx-auto my-8"></div>
        )}

        {filteredData && (
          <section>
            <div className="bg-blue-600 text-white rounded-t-lg p-5 text-center">
              <h2 className="text-4xl font-bold">
                របាយការណ៍អត្ថបទដែលត្រូវដាក់ចូលប្រព័ន្ធ
              </h2>
              {dateRange && <p className="text-2xl mt-2">{dateRange}</p>}
            </div>

            <div className="flex flex-col sm:flex-row justify-between items-center my-6 gap-4">
              <button
                onClick={downloadExcel}
                className="bg-green-600 text-white font-bold py-3 px-8 rounded-lg text-xl hover:bg-green-700 disabled:bg-gray-400 transition-all shadow-md hover:shadow-lg transform hover:-translate-y-0.5"
                disabled={!filteredData || filteredData.length === 0}
              >
                ទាញយក Excel
              </button>
            </div>

            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200 mb-6">
              <div className="flex flex-col sm:flex-row items-center gap-4">
                <label className="font-semibold text-gray-700 text-xl whitespace-nowrap">
                  ត្រងតាមអ្នកទស្សនា:
                </label>
                <div className="flex flex-wrap gap-3">
                  {VIEW_FILTERS.map((filter) => (
                    <button
                      key={filter.key}
                      onClick={() => setViewFilter(filter.key)}
                      className={`py-2 px-5 rounded-md text-lg font-semibold transition-colors ${
                        viewFilter === filter.key
                          ? "bg-blue-600 text-white shadow"
                          : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-100"
                      }`}
                    >
                      {filter.label}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <div
              className="overflow-x-auto max-h-[65vh] border border-gray-200 rounded-b-lg shadow-md"
              role="region"
              aria-labelledby="output-table-caption"
            >
              {filteredData.length > 0 ? (
                <table className="w-full text-lg text-left text-gray-800">
                  <caption id="output-table-caption" className="sr-only">
                    ទិន្នន័យដែលបានបញ្ចូល និងដំណើរការរួចរាល់សម្រាប់ទាញយក។
                  </caption>
                  <thead className="text-lg text-gray-700 uppercase bg-blue-100 sticky top-0">
                    <tr>
                      {outputHeaders.map((header) => (
                        <th
                          key={header}
                          scope="col"
                          className="px-6 py-6 font-bold"
                        >
                          <div className="flex items-center">
                            {header}
                            <svg
                              className="w-4 h-4 ml-2"
                              aria-hidden="true"
                              xmlns="http://www.w3.org/2000/svg"
                              fill="currentColor"
                              viewBox="0 0 24 24"
                            >
                              <path d="M8.574 11.024h6.852a2.075 2.075 0 0 0 1.847-1.086 1.9 1.9 0 0 0-.11-1.986L13.736 2.9a2.122 2.122 0 0 0-3.472 0L6.837 7.952a1.9 1.9 0 0 0-.11 1.986 2.074 2.074 0 0 0 1.847 1.086Zm6.852 1.952H8.574a2.075 2.075 0 0 0-1.847 1.087 1.9 1.9 0 0 0 .11 1.985l3.426 5.05a2.122 2.122 0 0 0 3.472 0l3.427-5.05a1.9 1.9 0 0 0 .11-1.985 2.074 2.074 0 0 0-1.847-1.087Z" />
                            </svg>
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, index) => (
                      <tr
                        key={row.id + "-" + index}
                        className="bg-white border-b hover:bg-gray-50"
                      >
                        {outputKeys.map((key) => (
                          <td key={key} className="px-6 py-6">
                            {row[key]}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="text-center p-16 text-gray-500 bg-white rounded-b-lg">
                  <p className="text-xl">
                    មិនមានទិន្នន័យត្រូវនឹងលក្ខខណ្ឌរបស់អ្នកទេ។
                  </p>
                </div>
              )}
            </div>
          </section>
        )}

        {!isLoading && !mergedData && (
          <div className="text-center text-gray-500 p-16 bg-gray-50 rounded-lg">
            <p className="text-2xl">
              ទិន្នន័យដែលបានដំណើរការនឹងបង្ហាញនៅទីនេះ បន្ទាប់ពីអ្នកផ្ទុកឡើង
              និងដំណើរការឯកសារ។
            </p>
          </div>
        )}
      </main>
    </div>
  );
};

const container = document.getElementById("root");
if (container) {
  const root = createRoot(container);
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}
