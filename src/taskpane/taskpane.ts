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

// Thêm biến để theo dõi số lượng sheet
let currentSheetCount = 0;

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
    // Format JSON với indent 2 spaces và thêm màu sắc
    const formattedJson = JSON.stringify(spec, null, 2)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(
        /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g,
        function (match) {
          let cls = "number";
          if (/^"/.test(match)) {
            if (/:$/.test(match)) {
              cls = "key";
            } else {
              cls = "string";
            }
          } else if (/true|false/.test(match)) {
            cls = "boolean";
          } else if (/null/.test(match)) {
            cls = "null";
          }
          return '<span class="' + cls + '">' + match + "</span>";
        }
      );

    specData.innerHTML = formattedJson;
  }

  const nameFormatDisplay: HTMLElement | null = document.getElementById("nameFormatDisplay");
  if (nameFormatDisplay) {
    nameFormatDisplay.textContent =
      spec.config.nameformat.length > 0 ? spec.config.nameformat.join("") : "Chưa có định dạng";
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

  const index: number = currentSheetCount++;
  const card: HTMLDivElement = document.createElement("div");
  card.className = "sheet-card";
  card.setAttribute("data-sheet-index", index.toString());

  const cardTemplate = `
    <button class="delete-btn" id="deleteBtn-${index}" title="Xóa sheet">×</button>
    
    <div class="sheet-header">
      <div class="input-group">
        <label>Sheet Name:</label>
        <input 
          type="text" 
          id="sheetName-${index}" 
          placeholder="Leave blank for active sheet" 
        />
      </div>
      <div class="input-group">
        <label>DB Table Name:</label>
        <input 
          type="text" 
          id="dbTableName-${index}" 
          placeholder="e.g., students" 
        />
      </div>
    </div>

    <div class="checkbox-container">
      <input type="checkbox" id="visible-${index}" checked />
      <label for="visible-${index}">Hiển thị sheet trong file kết quả</label>
    </div>

    <div class="form-action">
      <button onclick="saveSheet(${index})">Save</button>
    </div>

    <div class="mapping-section">
      <h4>Mapping</h4>
      <p>Selected Cell: <span class="selected-cell">None</span></p>
      
      <div class="form-row">
        <div class="input-group">
          <label>Type:</label>
          <select 
            id="mappingType-${index}" 
            onchange="toggleMappingInputs(${index})"
          >
            <option value="dbfield" selected>DB Field</option>
            <option value="dbfields">DB Fields (Combined)</option>
            <option value="const">Constant</option>
            <option value="extrafield">Extra Field</option>
            <option value="comment">Comment</option>
          </select>
        </div>
      </div>
      
      <div id="mappingInputs-${index}">
        <!-- DB Field Input -->
        <div id="dbfieldInput-${index}" style="display: block;">
          <div class="form-row">
            <div class="input-group">
              <label>DB Field:</label>
              <input 
                type="text" 
                id="dbfield-${index}" 
                placeholder="e.g., mssv" 
              />
            </div>
          </div>
        </div>
        
        <!-- DB Fields Combined Input -->
        <div id="dbfieldsInput-${index}" style="display: none;">
          <div class="form-row">
            <div class="input-group">
              <label>Format:</label>
              <input 
                type="text" 
                id="dbfieldsFormat-${index}" 
                placeholder="e.g., {0} {1} {2}" 
              />
            </div>
          </div>
          <div class="form-row">
            <div class="input-group">
              <label>Fields:</label>
              <input 
                type="text" 
                id="dbfieldsList-${index}" 
                placeholder="e.g., last_name, middle_name, first_name" 
              />
            </div>
          </div>
        </div>
        
        <!-- Constant Input -->
        <div id="constInput-${index}" style="display: none;">
          <div class="form-row">
            <div class="input-group">
              <label>Value:</label>
              <textarea 
                id="constValue-${index}" 
                placeholder="e.g., 2024.2"
              ></textarea>
            </div>
          </div>
        </div>
        
        <!-- Extra Field Input -->
        <div id="extrafieldInput-${index}" style="display: none;">
          <div class="form-row">
            <div class="input-group">
              <label>Extra Field:</label>
              <input 
                type="text" 
                id="extrafield-${index}" 
                placeholder="e.g., thesis_start_date" 
              />
            </div>
          </div>
        </div>
        
        <!-- Comment Input -->
        <div id="commentInput-${index}" style="display: none;">
          <div class="form-row">
            <div class="input-group">
              <label>Comment:</label>
              <input 
                type="text" 
                id="comment-${index}" 
                placeholder="e.g., Học kì" 
              />
            </div>
          </div>
        </div>
      </div>
      
      <div class="form-action">
        <button onclick="addMapping(${index})">Add Mapping</button>
      </div>
      
      <div class="mapping-list" id="mappingList-${index}"></div>
    </div>
  `;

  card.innerHTML = cardTemplate;

  sheetCards.appendChild(card);

  // Thêm event listener sau khi card đã được thêm vào DOM
  const deleteBtn = document.getElementById(`deleteBtn-${index}`);
  if (deleteBtn) {
    deleteBtn.addEventListener("click", function () {
      removeSheet(index);
    });
  }

  // Thêm sheet mới vào spec
  spec.sheets.push({
    name: "",
    visible: true,
    mapping: {
      dbtablename: "",
      cells: [],
    },
  });

  updateSelectedCells();
  updateSpecDisplay();
  updateMappingList(index);
}

// Lưu sheetName và dbTableName cho sheet cụ thể
function saveSheet(sheetIndex: number): void {
  if (sheetIndex >= spec.sheets.length) {
    console.error("Invalid sheet index:", sheetIndex);
    return;
  }

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

  console.log("Sheet saved:", spec.sheets[sheetIndex]);
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

// Cập nhật danh sách mapping cho sheet cụ thể
function updateMappingList(sheetIndex: number): void {
  const mappingList: HTMLElement | null = document.getElementById(`mappingList-${sheetIndex}`);
  if (!mappingList) return;

  mappingList.innerHTML = "";
  spec.sheets[sheetIndex].mapping.cells.forEach((mapping: MappingCell, mappingIndex: number) => {
    const item: HTMLDivElement = document.createElement("div");
    item.className = "mapping-item";
    let mappingText: string = `${mapping.cell}: `;
    if (mapping.dbfield) mappingText += `dbfield=${mapping.dbfield}`;
    else if (mapping.dbfields) mappingText += `dbfields=${mapping.dbfields.join(", ")}`;
    else if (mapping.const) mappingText += `const=${mapping.const}`;
    else if (mapping.extrafield) mappingText += `extrafield=${mapping.extrafield}`;
    else if (mapping.comment) mappingText += `comment=${mapping.comment}`;
    item.innerHTML = `
          <span>${mappingText}</span>
          <button onclick="deleteMapping(${sheetIndex}, ${mappingIndex})">Delete</button>
      `;
    mappingList.appendChild(item);
  });
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
        updateMappingList(sheetIndex);
        updateSpecDisplay();
      } else {
        alert("Please fill in the mapping details!");
      }
    });
  } catch (error) {
    console.error("Error adding mapping:", error);
  }
}

// Xóa mapping từ sheet cụ thể
function deleteMapping(sheetIndex: number, mappingIndex: number): void {
  // Xóa trực tiếp không cần xác nhận hoặc có thể giữ xác nhận nếu muốn
  spec.sheets[sheetIndex].mapping.cells.splice(mappingIndex, 1);
  updateMappingList(sheetIndex);
  updateSpecDisplay();
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

  // Hiển thị thông báo thành công
  const notification = document.createElement("div");
  notification.className = "success-notification";
  notification.textContent = "File đặc tả đã được tạo thành công!";
  document.body.appendChild(notification);

  setTimeout(() => {
    notification.remove();
  }, 3000);
}

// Hàm xử lý xóa sheet
function removeSheet(index: number): void {
  try {
    console.log("Removing sheet at index:", index);

    // Xóa khỏi spec
    spec.sheets.splice(index, 1);

    // Xóa card khỏi DOM
    const card = document.querySelector(`.sheet-card[data-sheet-index="${index}"]`);
    if (card) {
      card.remove();
    }

    // Giảm số lượng sheet
    currentSheetCount--;

    // Cập nhật lại chỉ số cho tất cả sheet card còn lại
    updateSheetIndexes();

    // Cập nhật hiển thị
    updateSpecDisplay();
  } catch (error) {
    console.error("Error removing sheet:", error);
  }
}

// Hàm mới để cập nhật lại chỉ số sau khi xóa sheet
function updateSheetIndexes(): void {
  const sheetCards = document.querySelectorAll(".sheet-card");

  sheetCards.forEach((card, newIndex) => {
    // Cập nhật thuộc tính data-sheet-index
    card.setAttribute("data-sheet-index", newIndex.toString());

    // Cập nhật ID của tất cả các phần tử trong card
    const elements = [
      "deleteBtn",
      "sheetName",
      "dbTableName",
      "visible",
      "mappingType",
      "mappingInputs",
      "dbfieldInput",
      "dbfieldsInput",
      "constInput",
      "extrafieldInput",
      "commentInput",
      "dbfield",
      "dbfieldsFormat",
      "dbfieldsList",
      "constValue",
      "extrafield",
      "comment",
      "mappingList",
    ];

    elements.forEach((prefix) => {
      const element = card.querySelector(`[id^="${prefix}-"]`);
      if (element) {
        element.id = `${prefix}-${newIndex}`;
      }
    });

    // Cập nhật các hàm onclick
    const buttons = card.querySelectorAll("button");
    buttons.forEach((button) => {
      const onclickAttr = button.getAttribute("onclick");
      if (onclickAttr) {
        // Thay thế chỉ số cũ bằng chỉ số mới trong chuỗi onclick
        const newOnclick = onclickAttr.replace(/\(\d+\)/, `(${newIndex})`);
        button.setAttribute("onclick", newOnclick);
      }
    });
  });
}

// Gắn các hàm vào window
(window as any).addNameFormat = addNameFormat;
(window as any).addSheetCard = addSheetCard;
(window as any).saveSheet = saveSheet;
(window as any).toggleMappingInputs = toggleMappingInputs;
(window as any).addMapping = addMapping;
(window as any).deleteMapping = deleteMapping;
(window as any).generateSpecFile = generateSpecFile;
(window as any).removeSheet = removeSheet;
