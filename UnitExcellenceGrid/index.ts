import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;

type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class UnitExcellenceGrid
    implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private container!: HTMLDivElement;
    private context!: ComponentFramework.Context<IInputs>;
    private dataset!: DataSet;

    private searchInput!: HTMLInputElement;
    private tableContainer!: HTMLDivElement;
    private statusElement!: HTMLDivElement;
    private previousButton!: HTMLButtonElement;
    private nextButton!: HTMLButtonElement;

    private baseFilter?: DataSetInterfaces.FilterExpression;

    private initialized = false;
    private pageSize = 5;
    private currentPage = 1;
    private pendingNextPage = false;

    private nativeToolbarObserver?: MutationObserver;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {

        this.context = context;
        this.container = container;

        this.container.style.fontFamily = "Segoe UI, Arial, sans-serif";
        this.container.style.width = "100%";

        this.createInterface();
        this.startNativeToolbarCleanup();
    }

    private createInterface(): void {

        const wrapper = document.createElement("div");

        wrapper.style.padding = "15px";
        wrapper.style.width = "100%";
        wrapper.style.boxSizing = "border-box";

        const searchSection = document.createElement("div");

        searchSection.style.display = "flex";
        searchSection.style.gap = "10px";
        searchSection.style.marginBottom = "20px";

        this.searchInput = document.createElement("input");

        this.searchInput.type = "text";
        this.searchInput.placeholder = "Search units...";
        this.searchInput.style.flex = "1";
        this.searchInput.style.padding = "12px";
        this.searchInput.style.border = "1px solid #ddd";
        this.searchInput.style.borderRadius = "5px";
        this.searchInput.style.fontSize = "14px";

        const searchButton = document.createElement("button");
        searchButton.textContent = "Search";
        this.styleButton(searchButton);

        const clearButton = document.createElement("button");
        clearButton.textContent = "Clear";
        this.styleButton(clearButton, "#666");

        searchButton.addEventListener("click", () => {
            this.applySearch();
        });

        this.searchInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                this.applySearch();
            }
        });

        clearButton.addEventListener("click", () => {
            this.searchInput.value = "";
            this.applySearch();
        });

        searchSection.appendChild(this.searchInput);
        searchSection.appendChild(searchButton);
        searchSection.appendChild(clearButton);

        this.statusElement = document.createElement("div");
        this.statusElement.style.marginBottom = "12px";
        this.statusElement.style.color = "#666";
        this.statusElement.style.fontSize = "13px";

        this.tableContainer = document.createElement("div");
        this.tableContainer.style.width = "100%";
        this.tableContainer.style.overflowX = "auto";

        const pagination = document.createElement("div");

        pagination.style.display = "flex";
        pagination.style.justifyContent = "flex-end";
        pagination.style.alignItems = "center";
        pagination.style.gap = "12px";
        pagination.style.marginTop = "15px";

        this.previousButton = document.createElement("button");
        this.previousButton.textContent = "Previous";
        this.styleButton(this.previousButton);

        this.nextButton = document.createElement("button");
        this.nextButton.textContent = "Next";
        this.styleButton(this.nextButton);

        this.previousButton.addEventListener("click", () => {

            if (this.currentPage > 1) {
                this.currentPage--;
                this.renderGrid();
            }

        });

        this.nextButton.addEventListener("click", () => {

            const allIds = this.dataset.sortedRecordIds;

            const nextPageStart = this.currentPage * this.pageSize;

            if (nextPageStart < allIds.length) {
                this.currentPage++;
                this.renderGrid();
                return;
            }

            if (this.dataset.paging.hasNextPage) {

                this.pendingNextPage = true;
                this.nextButton.disabled = true;

                this.statusElement.textContent = "Loading more units...";

                this.dataset.paging.loadNextPage();
            }

        });

        pagination.appendChild(this.previousButton);
        pagination.appendChild(this.nextButton);

        wrapper.appendChild(searchSection);
        wrapper.appendChild(this.statusElement);
        wrapper.appendChild(this.tableContainer);
        wrapper.appendChild(pagination);

        this.container.appendChild(wrapper);
    }

    private styleButton(
        button: HTMLButtonElement,
        background = "#106ebe"
    ): void {

        button.style.backgroundColor = background;
        button.style.color = "#fff";
        button.style.border = "none";
        button.style.padding = "11px 22px";
        button.style.borderRadius = "4px";
        button.style.cursor = "pointer";
        button.style.fontSize = "14px";
    }

    private hideNativeGridControls(): void {

        document
            .querySelectorAll<HTMLElement>(
                '[data-id="data-set-quickFind-container"]'
            )
            .forEach(element => {
                element.style.setProperty("display", "none", "important");
            });

        document
            .querySelectorAll<HTMLElement>(
                'i[data-icon-name="ColumnOptions"]'
            )
            .forEach(icon => {

                const button = icon.closest<HTMLElement>(
                    'button, [role="button"], .ms-Button'
                );

                const target = button || icon;

                target.style.setProperty("display", "none", "important");
            });
    }

    private startNativeToolbarCleanup(): void {

        this.hideNativeGridControls();

        this.nativeToolbarObserver = new MutationObserver(() => {
            this.hideNativeGridControls();
        });

        this.nativeToolbarObserver.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    public updateView(
        context: ComponentFramework.Context<IInputs>
    ): void {

        this.context = context;
        this.dataset = context.parameters.sampleDataSet;

        if (!this.initialized) {

            this.dataset.paging.setPageSize(this.pageSize);

            this.baseFilter = this.dataset.filtering.getFilter();

            this.initialized = true;

            this.setDatasetFilter("", true);

            return;
        }

        if (this.pendingNextPage && !this.dataset.loading) {

            const allIds = this.dataset.sortedRecordIds;
            const requiredRecords = this.currentPage * this.pageSize;

            if (allIds.length > requiredRecords) {
                this.currentPage++;
            }

            this.pendingNextPage = false;
        }

        this.renderGrid();
    }

    private setDatasetFilter(
        searchText: string,
        initial = false
    ): void {

        const searchableColumns = [
            "pw_name",
            "pw_buildingnumber"
        ];

        const searchValue = searchText.trim();

        const conditions: DataSetInterfaces.ConditionExpression[] =
            initial
                ? [{
                    attributeName: "pw_name",
                    conditionOperator: 0 as DataSetInterfaces.Types.ConditionOperator,
                    value: "__NO_UNIT_MATCH__"
                }]
                : [];

        const filters: DataSetInterfaces.FilterExpression[] = [];

        if (this.baseFilter) {
            filters.push(this.baseFilter);
        }

        if (searchValue) {

            const searchConditions: DataSetInterfaces.ConditionExpression[] =
                searchableColumns.map(column => ({
                    attributeName: column,
                    conditionOperator: 6 as DataSetInterfaces.Types.ConditionOperator,
                    value: `%${searchValue}%`
                }));

            filters.push({
                filterOperator: 1,
                conditions: searchConditions
            });
        }

        const filter: DataSetInterfaces.FilterExpression = {
            filterOperator: 0,
            conditions,
            filters
        };

        this.dataset.filtering.setFilter(filter);
        this.dataset.refresh();
    }

    private applySearch(): void {

        if (!this.dataset) {
            return;
        }

        this.currentPage = 1;
        this.pendingNextPage = false;

        const searchText = this.searchInput.value.trim();

        if (!searchText) {
            this.setDatasetFilter("", true);
            return;
        }

        this.setDatasetFilter(searchText);
    }

    private renderGrid(): void {

        if (this.dataset.loading) {
            this.statusElement.textContent = "Loading units...";
            return;
        }

        if (this.dataset.error) {
            this.statusElement.textContent =
                "Error loading units: " + this.dataset.errorMessage;
            return;
        }

        const allIds = this.dataset.sortedRecordIds;

        const startIndex = (this.currentPage - 1) * this.pageSize;

        const ids = allIds.slice(
            startIndex,
            startIndex + this.pageSize
        );

        this.tableContainer.innerHTML = "";

        if (ids.length === 0) {

            this.statusElement.textContent =
                "No units found. Search to display units.";

            this.previousButton.disabled = true;
            this.nextButton.disabled = true;

            return;
        }

        this.statusElement.textContent =
            `Showing ${startIndex + 1} - ${startIndex + ids.length} units`;

        const table = document.createElement("table");

        table.style.width = "100%";
        table.style.borderCollapse = "collapse";
        table.style.backgroundColor = "#fff";

        // Read visible columns dynamically from the selected Unit view
        const columns = this.dataset.columns.filter(
            column => !column.isHidden
        );

        const header = document.createElement("thead");
        const headerRow = document.createElement("tr");

        columns.forEach(column => {

            const th = document.createElement("th");

            th.textContent = column.displayName;
            th.style.padding = "13px";
            th.style.textAlign = "left";
            th.style.backgroundColor = "#f5f5f5";
            th.style.borderBottom = "1px solid #ddd";
            th.style.whiteSpace = "nowrap";

            headerRow.appendChild(th);
        });

        header.appendChild(headerRow);
        table.appendChild(header);

        const tbody = document.createElement("tbody");

        ids.forEach(id => {

            const record = this.dataset.records[id];

            if (!record) {
                return;
            }

            const row = document.createElement("tr");

            row.style.cursor = "pointer";

            columns.forEach(column => {

                const td = document.createElement("td");

                td.textContent =
                    record.getFormattedValue(column.name) || "";

                td.style.padding = "13px";
                td.style.borderBottom = "1px solid #eee";
                td.style.whiteSpace = "nowrap";

                row.appendChild(td);
            });

            row.addEventListener("dblclick", () => {

                this.context.navigation.openForm({
                    entityName: "pw_unit",
                    entityId: record.getRecordId()
                });

            });

            tbody.appendChild(row);
        });

        table.appendChild(tbody);
        this.tableContainer.appendChild(table);

        this.previousButton.disabled = this.currentPage === 1;

        this.nextButton.disabled =
            (this.currentPage * this.pageSize >= allIds.length) &&
            !this.dataset.paging.hasNextPage;
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {

        if (this.nativeToolbarObserver) {
            this.nativeToolbarObserver.disconnect();
            this.nativeToolbarObserver = undefined;
        }

        this.container.innerHTML = "";
    }
}