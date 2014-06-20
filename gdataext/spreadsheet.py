import math
import gdata.spreadsheets.data
import gdata.spreadsheet.service
import gdataext


class ClientConfig(object):
    SERVICE = gdata.spreadsheet.service.SpreadsheetsService
    APP_NAME = 'gdata_spreadsheep_api'
    DEBUG = False


def create_client(**kwa):
    """
    """
    config = ClientConfig()
    for k, v in kwa.items():
        setattr(config, k, v)
    return gdataext.create_client(config)


class Worksheet(object):
    def __init__(self, client, sheet_id, wsheet_id):
        """Construct a `Worksheet` object with the given `client`,
        `sheet_id` and 'wsheet_id'.
        """
        self._client = client
        self._sheet_id = sheet_id
        self._wsheet_id = wsheet_id
        self._sheet = gdataext.search_feed_id(self._sheet_id,
                                      self._client.GetSpreadsheetsFeed())
        self._wsheet = gdataext.search_feed_id(self._wsheet_id,
                                       self._client.GetWorksheetsFeed(self._sheet_id))

    def _iter_rows(self):
        """Iterate the rows in this worksheet.
        """
        for row in self._client.GetListFeed(self._sheet_id,
                                            wksht_id=self._wsheet_id).entry:
            yield row

    def get_row(self, row_idx):
        """Get a row by it's index.
        # Not in use.
        """
        i = 0
        for row in self._iter_rows():
            if i == row_idx:
                return row
            i += 1

    def update_sheet(self):
        self._wsheet = self._client.UpdateWorksheet(self._wsheet)

    @property
    def row_count(self):
        return int(self._wsheet.row_count.text)

    @row_count.setter
    def row_count(self, val):
        self._wsheet.row_count.text = str(val)
        self.update_sheet()

    @property
    def cells_feed(self):
        return self._client.GetCellsFeed(self._sheet_id, self._wsheet_id)

    def start_batch(self):
        #?
        return gdata.spreadsheet.SpreadsheetsCellsFeed()

    def execute_batch(self, batch, cells_feed=None):
        cells_feed = cells_feed or self.cells_feed
        return self._client.ExecuteBatch(batch, cells_feed.GetBatchLink().href)

    def clear(self):
        cells_feed = self.cells_feed
        batch = self.start_batch()
        self.batch_clear(batch, cells_feed)
        return self._client.ExecuteBatch(batch, cells_feed.GetBatchLink().href)

    def batch_cells_feed(self, cols=2, rows=100, col_offset=0, row_offset=0):
        query = gdata.spreadsheet.service.CellQuery()
        query.return_empty = "true"
        query.min_row = str(1 + row_offset)

        max_row = rows + row_offset

        query.min_col = str(1 + col_offset)
        query.max_col = str(cols + col_offset)

        if max_row > self.row_count:
            query.max_row = str(self.row_count)
        else:
            query.max_row = str(max_row)
            

        print(query.min_row, query.max_row, query.min_col, query.max_col)

        return self._client.GetCellsFeed(self._sheet_id,
                    wksht_id=self._wsheet_id, query=query)

    def batch_update_row(self, batch, cells_feed, row, *vals):
        head_cell = cells_feed.entry[0].cell
        head_col = head_cell.col
        head_row = head_cell.row

        # Get max columns.
        for entry in cells_feed.entry:
            if entry.cell.row != head_row:
                break
            else:
                max_col = int(entry.cell.col)

        assert max_col == len(vals), "Number of columns (%s) in cell_feed " \
                "does not match the number of vals (%s) being updated." % \
                    (max_col, len(vals))

        start_idx = max_col * (row-1)
        end_idx = max_col * (row)

        for i, entry in enumerate(cells_feed.entry[start_idx:end_idx]):
            entry.cell.inputValue = str(vals[i])
            batch.AddUpdate(entry)

        return batch

    def batch_clear(self, batch, cells_feed):
        for i, entry in enumerate(cells_feed.entry):
            entry.cell.inputValue = ''
            batch.AddUpdate(entry)

    def batch_add_rows(self, rows, batch_rows=100, row_offset=0):
        if not rows:
            return
        # Resize sheet if needed.
        if self.row_count < len(rows) + row_offset:
            self.row_count = len(rows) + row_offset

        max_col = len(rows[0])

        for i, row in enumerate(rows):
            if i % batch_rows == 0:
                p = 0
                batch = self.start_batch()
                cells_feed = self.batch_cells_feed(cols=max_col, rows=batch_rows,
                                                   row_offset=i)
            p += 1
            self.batch_update_row(batch, cells_feed, p, *row)

            if i % batch_rows == batch_rows - 1:
                self.execute_batch(batch)

        self.execute_batch(batch)

    def update_title(self, title):
        """
        """
        self._wsheet = gdataext.search_feed_id(self._wsheet_id,
                                       self._client.GetWorksheetsFeed(self._sheet_id))
        self._wsheet.title.text = title
        self._wsheet = self._client.UpdateWorksheet(self._wsheet)

    def update_cell(self, row, col, val):
        """Update a cell with the given row and col position. (1 index)
        """
        return self._client.UpdateCell(row, col, val, self._sheet_id,
                                       wksht_id=self._wsheet_id)

    def insert_row(self, data):
        """Inserts a row in to the spreadsheet.
        """
        return self._client.InsertRow(data, self._sheet_id, self._wsheet_id)