// Định nghĩa kiểu dữ liệu cho file đặc tả
interface MappingCell {
  cell: string;
  dbfield?: string;
  dbfields?: string[];
  const?: string;
  extrafield?: string;
  comment?: string;
}

interface Mapping {
  dbtablename: string;
  cells: MappingCell[];
}

interface Sheet {
  name: string;
  visible: boolean;
  mapping: Mapping;
}

interface Config {
  nameformat: string[];
}

interface SpecFile {
  config: Config;
  sheets: Sheet[];
  errMessage: string;
}

// Khai báo biến toàn cục
let spec: SpecFile = {
  config: { nameformat: [] },
  sheets: [],
  errMessage: "",
};

// Khởi tạo khi add-in được tải
Office.onReady(() => {
  updateSelectedCells().catch(console.error);
  Excel.run(async (context: Excel.RequestContext) => {
    const workbook: Excel.Workbook = context.workbook;
    workbook.onSelectionChanged.add(handleSelectionChange);
    await context.sync();
  }).catch(console.error);
  updateSpecDisplay(); // Hiển thị dữ liệu ban đầu
});

// Xử lý khi ô được chọn thay đổi
async function handleSelectionChange(): Promise<void> {
  await updateSelectedCells();
}

// Cập nhật ô đang chọn cho tất cả card
async function updateSelectedCells(): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      const cellAddress: string = range.address.split("!")[1] || "None";
      document.querySelectorAll(".selected-cell").forEach((el: Element) => {
        (el as HTMLElement).textContent = cellAddress;
      });
    });
  } catch (error) {
    console.error("Error updating selected cells:", error);
  }
}

// Cập nhật hiển thị dữ liệu spec
function updateSpecDisplay(): void {
  const specData: HTMLElement | null = document.getElementById("spec-data");
  if (specData) {
    specData.textContent = JSON.stringify(spec, null, 2);
  }
}

// Thêm nameFormat vào config
function addNameFormat(): void {
  const nameFormatInput: HTMLInputElement | null = document.getElementById(
    "nameFormat"
  ) as HTMLInputElement;
  const nameFormat: string = nameFormatInput?.value.trim() ?? "";
  if (nameFormat) {
    spec.config.nameformat = nameFormat.split(/(?=\?)/);
    if (nameFormatInput) nameFormatInput.value = "";
    console.log("Config updated:", spec.config);
    updateSpecDisplay();
  } else {
    alert("Please enter a name format!");
  }
}

// Thêm card nhập liệu cho sheet mới
function addSheetCard(): void {
  const sheetCards: HTMLElement | null = document.getElementById("sheetCards");
  if (!sheetCards) return;

  const index: number = spec.sheets.length;
  const card: HTMLDivElement = document.createElement("div");
  card.className = "sheet-card";
  card.innerHTML = `
      <label>Sheet Name:</label>
      <input type="text" id="sheetName-${index}" placeholder="Leave blank for active sheet" />
      <label>DB Table Name:</label>
      <input type="text" id="dbTableName-${index}" placeholder="e.g., students" />
      <label><input type="checkbox" id="visible-${index}" checked /> Visible</label>

      <div class="mapping-section">
          <h4>Mapping</h4>
          <p>Selected Cell: <span class="selected-cell">None</span></p>
          <label>Type:</label>
          <select id="mappingType-${index}" onchange="toggleMappingInputs(${index})">
              <option value="dbfield">DB Field</option>
              <option value="dbfields">DB Fields (Combined)</option>
              <option value="const">Constant</option>
              <option value="extrafield">Extra Field</option>
              <option value="comment">Comment</option>
          </select>
          <div id="mappingInputs-${index}">
              <div id="dbfieldInput-${index}" style="display: none;">
                  <label>DB Field:</label>
                  <input type="text" id="dbfield-${index}" placeholder="e.g., mssv" />
              </div>
              <div id="dbfieldsInput-${index}" style="display: none;">
                  <label>Format:</label>
                  <input type="text" id="dbfieldsFormat-${index}" placeholder="e.g., {0} {1} {2}" />
                  <label>Fields:</label>
                  <input type="text" id="dbfieldsList-${index}" placeholder="e.g., last_name, middle_name, first_name" />
              </div>
              <div id="constInput-${index}" style="display: none;">
                  <label>Value:</label>
                  <textarea id="constValue-${index}" placeholder="e.g., 2024.2"></textarea>
              </div>
              <div id="extrafieldInput-${index}" style="display: none;">
                  <label>Extra Field:</label>
                  <input type="text" id="extrafield-${index}" placeholder="e.g., thesis_start_date" />
              </div>
              <div id="commentInput-${index}" style="display: none;">
                  <label>Comment:</label>
                  <input type="text" id="comment-${index}" placeholder="e.g., Học kì" />
              </div>
          </div>
          <button onclick="addMapping(${index})">Add Mapping</button>
      </div>
  `;
  sheetCards.appendChild(card);

  // Thêm sheet rỗng vào spec
  spec.sheets.push({
    name: "",
    visible: true,
    mapping: {
      dbtablename: "",
      cells: [],
    },
  });
  updateSelectedCells(); // Cập nhật ô được chọn cho card mới
  updateSpecDisplay();
}

// Hiển thị input tương ứng với mapping type cho sheet cụ thể
function toggleMappingInputs(sheetIndex: number): void {
  const mappingTypeSelect: HTMLSelectElement | null = document.getElementById(
    `mappingType-${sheetIndex}`
  ) as HTMLSelectElement;
  const mappingType: string = mappingTypeSelect?.value ?? "dbfield";
  const inputs: string[] = [
    "dbfieldInput",
    "dbfieldsInput",
    "constInput",
    "extrafieldInput",
    "commentInput",
  ];

  inputs.forEach((id: string) => {
    const element: HTMLElement | null = document.getElementById(`${id}-${sheetIndex}`);
    if (element) element.style.display = "none";
  });

  const activeInput: HTMLElement | null = document.getElementById(
    `${mappingType}Input-${sheetIndex}`
  );
  if (activeInput) activeInput.style.display = "block";
}

// Thêm mapping cho sheet cụ thể
async function addMapping(sheetIndex: number): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      const cellAddress: string = range.address.split("!")[1];

      // Cập nhật sheet name và dbTableName từ input
      const sheetNameInput: HTMLInputElement | null = document.getElementById(
        `sheetName-${sheetIndex}`
      ) as HTMLInputElement;
      const dbTableNameInput: HTMLInputElement | null = document.getElementById(
        `dbTableName-${sheetIndex}`
      ) as HTMLInputElement;
      const visibleInput: HTMLInputElement | null = document.getElementById(
        `visible-${sheetIndex}`
      ) as HTMLInputElement;

      spec.sheets[sheetIndex].name = sheetNameInput?.value.trim() ?? "";
      spec.sheets[sheetIndex].mapping.dbtablename = dbTableNameInput?.value.trim() ?? "";
      spec.sheets[sheetIndex].visible = visibleInput?.checked ?? true;

      if (!spec.sheets[sheetIndex].mapping.dbtablename) {
        alert("Please enter a DB Table Name for this sheet!");
        return;
      }

      const mappingTypeSelect: HTMLSelectElement | null = document.getElementById(
        `mappingType-${sheetIndex}`
      ) as HTMLSelectElement;
      const mappingType: string = mappingTypeSelect?.value ?? "dbfield";
      const mapping: MappingCell = { cell: cellAddress };

      switch (mappingType) {
        case "dbfield":
          const dbfieldInput: HTMLInputElement | null = document.getElementById(
            `dbfield-${sheetIndex}`
          ) as HTMLInputElement;
          const dbfield: string = dbfieldInput?.value.trim() ?? "";
          if (dbfield) mapping.dbfield = dbfield;
          break;
        case "dbfields":
          const formatInput: HTMLInputElement | null = document.getElementById(
            `dbfieldsFormat-${sheetIndex}`
          ) as HTMLInputElement;
          const fieldsInput: HTMLInputElement | null = document.getElementById(
            `dbfieldsList-${sheetIndex}`
          ) as HTMLInputElement;
          const format: string = formatInput?.value.trim() ?? "";
          const fields: string[] = fieldsInput?.value.split(",").map((f: string) => f.trim()) ?? [];
          if (format && fields.length) mapping.dbfields = [format, ...fields];
          break;
        case "const":
          const constInput: HTMLTextAreaElement | null = document.getElementById(
            `constValue-${sheetIndex}`
          ) as HTMLTextAreaElement;
          const constValue: string = constInput?.value.trim() ?? "";
          if (constValue) mapping.const = constValue;
          break;
        case "extrafield":
          const extrafieldInput: HTMLInputElement | null = document.getElementById(
            `extrafield-${sheetIndex}`
          ) as HTMLInputElement;
          const extrafield: string = extrafieldInput?.value.trim() ?? "";
          if (extrafield) mapping.extrafield = extrafield;
          break;
        case "comment":
          const commentInput: HTMLInputElement | null = document.getElementById(
            `comment-${sheetIndex}`
          ) as HTMLInputElement;
          const comment: string = commentInput?.value.trim() ?? "";
          if (comment) mapping.comment = comment;
          break;
      }

      if (Object.keys(mapping).length > 1) {
        spec.sheets[sheetIndex].mapping.cells.push(mapping);
        clearMappingInputs(sheetIndex);
        console.log("Mapping added to sheet", sheetIndex, ":", mapping);
        updateSpecDisplay();
      } else {
        alert("Please fill in the mapping details!");
      }
    });
  } catch (error) {
    console.error("Error adding mapping:", error);
  }
}

// Xóa input sau khi thêm mapping cho sheet cụ thể
function clearMappingInputs(sheetIndex: number): void {
  const inputs: string[] = [
    `dbfield-${sheetIndex}`,
    `dbfieldsFormat-${sheetIndex}`,
    `dbfieldsList-${sheetIndex}`,
    `constValue-${sheetIndex}`,
    `extrafield-${sheetIndex}`,
    `comment-${sheetIndex}`,
  ];
  inputs.forEach((id: string) => {
    const element: HTMLInputElement | HTMLTextAreaElement | null = document.getElementById(id) as
      | HTMLInputElement
      | HTMLTextAreaElement;
    if (element) element.value = "";
  });
}

// Sinh file đặc tả
function generateSpecFile(): void {
  // Cập nhật dữ liệu từ tất cả card trước khi sinh file
  spec.sheets.forEach((_: Sheet, index: number) => {
    const sheetNameInput: HTMLInputElement | null = document.getElementById(
      `sheetName-${index}`
    ) as HTMLInputElement;
    const dbTableNameInput: HTMLInputElement | null = document.getElementById(
      `dbTableName-${index}`
    ) as HTMLInputElement;
    const visibleInput: HTMLInputElement | null = document.getElementById(
      `visible-${index}`
    ) as HTMLInputElement;

    spec.sheets[index].name = sheetNameInput?.value.trim() ?? "";
    spec.sheets[index].mapping.dbtablename = dbTableNameInput?.value.trim() ?? "";
    spec.sheets[index].visible = visibleInput?.checked ?? true;
  });

  if (spec.sheets.length === 0 && spec.config.nameformat.length === 0) {
    alert("Please add some data to generate the spec file!");
    return;
  }

  const jsonString: string = JSON.stringify(spec, null, 2);
  const blob: Blob = new Blob([jsonString], { type: "application/json" });
  const link: HTMLAnchorElement = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "spec.json";
  link.click();
  console.log("Spec file generated:", spec);
}

// Gắn các hàm vào window để HTML có thể gọi
(window as any).addNameFormat = addNameFormat;
(window as any).addSheetCard = addSheetCard;
(window as any).toggleMappingInputs = toggleMappingInputs;
(window as any).addMapping = addMapping;
(window as any).generateSpecFile = generateSpecFile;
