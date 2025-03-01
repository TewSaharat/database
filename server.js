
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const WebSocket = require('ws');
const ExcelJS = require('exceljs');

// Initialize Express app
const app = express();
const port = 3000;

// Enable CORS0
app.use(cors());
app.use(express.json());

// Connect to SQLite database
const db = new sqlite3.Database('./route_db/route_db.db', (err) => {
  if (err) {
    console.error('Failed to connect to SQLite database:', err.message);
  } else {
    console.log('Connected to SQLite database.');
  }
});

// Set up WebSocket server
const wss = new WebSocket.Server({ server: app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
})});

// Broadcast updates to all connected clients
function broadcastUpdate(message) {
  wss.clients.forEach(client => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(JSON.stringify(message));
    }
  });
}

// API: Fetch routes with status = 0
app.get('/api/get-routes', (req, res) => {
  const query = `
    SELECT 
      cat_id, lamp_type, dir, dir_num, routes, control, km, 
      lat, long AS longitude, fovy, range, name_id, status, complaintReason, report_time
    FROM routes
    WHERE status = 0
  `;

  db.all(query, [], (err, rows) => {
    if (err) {
      console.error('Error fetching data:', err.message);
      res.status(500).json({ error: 'Failed to fetch data' });
    } else {
      res.json(rows);
    }
  });
});

// API: Fetch routes with filters
app.get('/api/routes', (req, res) => {
  const { category, routes } = req.query;
  let query = `
    SELECT 
      lat, long AS lng, name_id, routes, cat_id AS category, status, complaintReason, report_time
    FROM routes
  `;
  const conditions = [];
  if (category && category !== 'all') {
    conditions.push(`cat_id = '${category}'`);
  }
  if (routes && routes !== 'all') {
    conditions.push(`routes = '${routes}'`);
  }
  if (conditions.length > 0) {
    query += ' AND ' + conditions.join(' AND ');
  }
  console.log('Executing Query:', query);
  db.all(query, [], (err, rows) => {
    if (err) {
      console.error('Database Error:', err.message);
      res.status(500).json({ error: err.message });
    } else {
      res.json(rows);
    }
  });
});


// API: Update electric pole information and save to 'notify' and 'Repair_completed'
app.post('/api/save-electric-pole', (req, res) => {
  console.log('Request Body:', req.body);

  const {
    name_id,
    lampType,
    controller_edit,
    constructionDate,
    contractNumber,
    notes,
    status,
    repairMethod,
    complaintChannel,
    complaintCode,
    complaintTopic,
    complaintReason,
    lastRepairDate,
    controlType,
    repairItems,
    report_time
  } = req.body;

  if (!name_id) {
    return res.status(400).json({ error: 'Missing required field: name_id' });
  }
  const statusValue = status === true || status === 1 ? 1 : 0;
  // SQL query to update 'routes'
  const updateQuery = `
    UPDATE routes SET
      lampType_edit = ?, controller_edit = ?, constructionDate = ?, contractNumber = ?, notes = ?,
      status = ?, repairMethod = ?, complaintChannel = ?, complaintCode = ?, complaintTopic = ?,
      complaintReason = ?, lastRepairDate = ?, controlType = ?, repairItems = ?, report_time = ?
    WHERE name_id = ?
  `;
  db.run(updateQuery, [
    lampType, controller_edit, constructionDate, contractNumber, notes,
    statusValue, repairMethod, complaintChannel, complaintCode, complaintTopic,
    complaintReason, lastRepairDate, controlType, JSON.stringify(repairItems),
    report_time, name_id
  ], function (err) {
    if (err) {
      console.error('Update Error:', err.message);
      return res.status(500).json({ error: 'Failed to update data: ' + err.message });
    }
    console.log('Routes updated successfully.');
    // Fetch data from 'routes' for the given name_id
    const selectQuery = `SELECT * FROM routes WHERE name_id = ?`;
    db.get(selectQuery, [name_id], (err, row) => {
      if (err) {
        console.error('Fetch Error:', err.message);
        return res.status(500).json({ error: 'Failed to fetch data: ' + err.message });
      }
      if (!row) {
        return res.status(404).json({ error: `No data found for name_id: ${name_id}` });
      }
      console.log('Fetched Data from routes:', row);
      const now = new Date();
      const time = `${now.getDate().toString().padStart(2, '0')}-${(now.getMonth() + 1)
        .toString()
        .padStart(2, '0')}-${now.getFullYear()} ${now.getHours().toString().padStart(2, '0')}:${now
        .getMinutes()
        .toString()
        .padStart(2, '0')}`;

        
        if (statusValue === 1) {
      // Insert data into 'Repair_completed'
      const insertRepairCompletedQuery = `
        INSERT INTO Repair_completed (
          name_id, lampType, repairMethod, lastRepairDate, notes, repairItems,
          cat_id, dir, dir_num, routes, control, km, lat, long, fovy, range, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;

      db.run(insertRepairCompletedQuery, [
        row.name_id,
        row.lamp_type,
        row.repairMethod,
        row.lastRepairDate || time,
        row.notes ,
        row.repairItems || '{}',
        row.cat_id,
        row.dir,
        row.dir_num,
        row.routes,
        row.control,
        row.km,
        row.lat,
        row.long,
        row.fovy,
        row.range,
        row.status
      ], function (err) {
        if (err) {
          console.error('Insert Error (Repair_completed):', err.message);
          return res.status(500).json({ error: 'Failed to insert data into Repair_completed: ' + err.message });
        }

        console.log('Data inserted into Repair_completed successfully.');
        res.json({ message: 'Data saved successfully in Repair_completed.', changes: this.changes });
      });}else{
        // Insert data into 'notify'
        const insertNotifyQuery = `
          INSERT INTO notify (
            lamp_type, dir, dir_num, routes, control, km, lat, long, fovy, range,
            name_id, status, complaintReason, report_time
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;

        db.run(insertNotifyQuery, [
          row.lamp_type, row.dir, row.dir_num, row.routes, row.control, row.km,
          row.lat, row.long, row.fovy, row.range,
          name_id, statusValue, row.complaintReason || 'No complaint', time
        ], function (err) {
          if (err) {
            console.error('Insert Error (notify):', err.message);
            return res.status(500).json({ error: 'Failed to insert data into notify: ' + err.message });
          }

          console.log('Data inserted into notify successfully.');
          res.json({
            message: 'Data saved successfully in Repair_completed and notify.',
            changes: this.changes });
        });
      }

    });
   });
});

// API: Get marker by name_id
app.get('/api/marker/:name_id', (req, res) => {
  const name_id = decodeURIComponent(req.params.name_id);

  if (!name_id) {
    return res.status(400).json({ error: 'name_id is required' });
  }

  const query = `SELECT * FROM routes WHERE name_id = ?`;

  db.get(query, [name_id], (err, row) => {
    if (err) {
      console.error('Database Error:', err.message);
      return res.status(500).json({ error: 'Failed to fetch marker data' });
    }

    if (!row) {
      return res.status(404).json({ error: 'Marker not found' });
    }

    res.json(row);
  });
});

// API: Export notify table to Excel
app.get('/api/export-notify-to-excel', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Notify');

    // Define columns for the worksheet
    sheet.columns = [
      { header: 'Lamp Type', key: 'lamp_type', width: 15 },
      { header: 'Direction', key: 'dir', width: 10 },
      { header: 'Direction Number', key: 'dir_num', width: 15 },
      { header: 'Routes', key: 'routes', width: 15 },
      { header: 'Control', key: 'control', width: 10 },
      { header: 'KM', key: 'km', width: 15 },
      { header: 'Latitude', key: 'lat', width: 20 },
      { header: 'Longitude', key: 'long', width: 20 },
      { header: 'Field of View', key: 'fovy', width: 15 },
      { header: 'Range', key: 'range', width: 15 },
      { header: 'Name ID', key: 'name_id', width: 25 },
      { header: 'Status', key: 'status', width: 10 },
      { header: 'Lamp Type Edit', key: 'lampType_edit', width: 15 },
      { header: 'Controller Edit', key: 'controller_edit', width: 20 },
      { header: 'Construction Date', key: 'constructionDate', width: 20 },
      { header: 'Contract Number', key: 'contractNumber', width: 20 },
      { header: 'Repair Method', key: 'repairMethod', width: 20 },
      { header: 'Complaint Channel', key: 'complaintChannel', width: 20 },
      { header: 'Complaint Code', key: 'complaintCode', width: 15 },
      { header: 'Complaint Topic', key: 'complaintTopic', width: 20 },
      { header: 'Complaint Reason', key: 'complaintReason', width: 25 },
      { header: 'Repair Items', key: 'repairItems', width: 25 },
      { header: 'Control Type', key: 'controlType', width: 20 },
      { header: 'Last Repair Date', key: 'lastRepairDate', width: 20 },
      { header: 'Report Time', key: 'report_time', width: 20 },
    ];

    // Fetch data from the notify table
    const notifyData = await new Promise((resolve, reject) => {
      db.all('SELECT * FROM notify', [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });

    // Add rows to the worksheet
    sheet.addRows(notifyData);

    // Send the file as a download
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader('Content-Disposition', 'attachment; filename="notify_data.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error exporting notify data:', error.message);
    res.status(500).send('Error exporting notify data.');
  }
});

app.get('/api/export-repair-to-excel', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Repair Completed');

    // กำหนดคอลัมน์ในไฟล์ Excel ให้ตรงกับ notify
    sheet.columns = [
      { header: 'Lamp Type', key: 'lamp_type', width: 15 },
      { header: 'Direction', key: 'dir', width: 10 },
      { header: 'Direction Number', key: 'dir_num', width: 15 },
      { header: 'Routes', key: 'routes', width: 15 },
      { header: 'Control', key: 'control', width: 10 },
      { header: 'KM', key: 'km', width: 15 },
      { header: 'Latitude', key: 'lat', width: 20 },
      { header: 'Longitude', key: 'long', width: 20 },
      { header: 'Field of View', key: 'fovy', width: 15 },
      { header: 'Range', key: 'range', width: 15 },
      { header: 'Name ID', key: 'name_id', width: 25 },
      { header: 'Status', key: 'status', width: 10 },
      { header: 'Lamp Type Edit', key: 'lampType_edit', width: 15 },
      { header: 'Controller Edit', key: 'controller_edit', width: 20 },
      { header: 'Construction Date', key: 'constructionDate', width: 20 },
      { header: 'Contract Number', key: 'contractNumber', width: 20 },
    ];

    // ดึงข้อมูลทั้งหมดจากตาราง Repair_completed
    const repairData = await new Promise((resolve, reject) => {
      db.all('SELECT * FROM Repair_completed', [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });

    // เพิ่มข้อมูลในแผ่นงาน
    sheet.addRows(repairData);

    // ส่งไฟล์เป็นดาวน์โหลด
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader('Content-Disposition', 'attachment; filename="repair_completed.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error exporting repair data:', error.message);
    res.status(500).send('Error exporting repair data.');
  }
});
