using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using CMS.Models;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace CMS.Data
{
    public class CategoryField
    {
        public string Key   { get; set; } = "";
        public string Label { get; set; } = "";
    }

    public static class ComponentRepository
    {
        
        public static string ConnectionString { get; set; } =
            "Server=192.10.8.120,1433;Database=AltiumPCB;User Id=PCBadmin;Password=Sa112233;TrustServerCertificate=True;";

        private static List<Component> _components = new();
        private static int _nextId = 1;

        public static Dictionary<string, List<(string Key, string Label)>> CategoryFields = new()
        {
            ["RES"]     = new() { ("Resistance(Ohm)", "Resistance"), ("Power(W)", "Power"), ("Tolerance(%)", "Tolerance") },
            ["CAP"]     = new() { ("Capacitance(F)", "Capacitance"), ("Voltage(V)", "Voltage"), ("Tolerance(%)", "Tolerance"), ("Temperature Coefficient", "Temp Coeff"), ("Failure Mode", "Failure Mode") },
            ["IND"]     = new() { ("Inductance(H)", "Inductance"), ("Irms Current(A)", "Irms Current"), ("Tolerance(%)", "Tolerance"), ("Max DCR(Ohm)", "Max DCR") },
            ["FB"]      = new() { ("Impedance(Ohm)@Frequency(Hz)", "Impedance@Freq"), ("Max Rated Current(A)", "Max Current"), ("Max DCR(Ohm)", "Max DCR") },
            ["CMC"]     = new() { ("Inductance(H)@Frequency(Hz)", "Inductance@Freq"), ("Impedance(Ohm)@Frequency(Hz)", "Impedance@Freq"), ("Max Rated Current(A)", "Max Current"), ("Max DCR(Ohm)", "Max DCR") },
            ["DIO-REC"] = new() { ("Max Peak Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Max Forward Current(A)", "Forward Current") },
            ["DIO-SW"]  = new() { ("Max Peak Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Max Forward Current(A)", "Forward Current") },
            ["DIO-ZEN"] = new() { ("Nom Zener Voltage(V)", "Zener Voltage"), ("Tolerance(%)", "Tolerance") },
            ["DIO-SCH"] = new() { ("Max Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Avg Forward Current(V)", "Avg Current") },
            ["DIO-TVS"] = new() { ("Reverse Standoff Voltage(V)", "Standoff Voltage"), ("Peak Pulse Power(W)", "Peak Power"), ("Max Peak Pulse Current(A)", "Peak Current"), ("Polarization", "Polarization") },
            ["NTC"]     = new() { ("Resistance@25°C(Ohm)", "R@25°C"), ("Resistance Tolerance(%)", "Tolerance"), ("Beta(K)", "Beta"), ("Beta Tolerance(%)", "Beta Tol.") },
            ["PTC"]     = new() { ("Resistance@25°C(Ohm)", "R@25°C"), ("Percentage R. Tolerance(%)", "Tolerance") },
            ["LED"]     = new() { ("Nom Forward Voltage(V)", "Forward Voltage"), ("Nom Forward Current(A)", "Forward Current"), ("BIN", "BIN") },
            ["TRAN"]    = new() { ("Type", "Type (NPN/PNP)"), ("Max Collector-Emitter Voltage(V)", "Vce Max"), ("Max Collector Current(I)", "Ic Max"), ("Max Base-Emitter Voltage(V)", "Vbe Max"), ("Power(W)", "Power") },
            ["MOS"]     = new() { ("Type", "Type (N-Ch/P-Ch)"), ("Max Drain-Source Voltage(V)", "Vds Max"), ("Max Drain Current(I)", "Id Max"), ("Max Gate-Source Voltage(V)", "Vgs Max"), ("Power(W)", "Power") },
            ["DC-DC"]   = new() { ("Type", "Converter Type"), ("Max Output Voltage(V)", "Output Voltage"), ("Max Output Current(I)", "Output Current") },
            ["LDO"]     = new() { ("Output Type", "Output Type"), ("Max Output Voltage(V)", "Output Voltage"), ("Max Output Current(I)", "Output Current") },
            ["LED-DRV"] = new() { ("Type", "Driver Type"), ("Total Channel Count", "Channels"), ("Output Current/Channel(A)", "Current/Ch") },
            ["MCU"]     = new() { ("Core Processor", "Core"), ("Core Size(Bit)", "Core Size"), ("Clock Speed(Hz)", "Clock Speed"), ("Program Memory Size(B)", "Flash Size"), ("Program Memory Type", "Memory Type"), ("ASIL Category", "ASIL") },
            ["OTHER-IC"]= new() { },
            ["CON"]     = new() { ("Gender", "Gender"), ("Pin Count", "Pin Count"), ("Pitch(mm)", "Pitch"), ("Color", "Color") },
            ["TERM"]    = new() { ("Gender", "Gender"), ("Base Material", "Base Material"), ("Plating Material", "Plating") },
            ["PAD"]     = new() { ("PAD X-Size(mm)", "PAD X"), ("PAD Y-Size(mm)", "PAD Y"), ("Hole X-Size(mm)", "Hole X"), ("Hole Y-Size(mm)", "Hole Y"), ("PAD Shape", "PAD Shape"), ("Hole Shape", "Hole Shape") },
        };

        private static readonly Dictionary<string, List<(string Key, string Label)>> _builtInCategories =
            new(CategoryFields.ToDictionary(k => k.Key, k => new List<(string, string)>(k.Value)));

        public static List<string> AllCategories => CategoryFields.Keys.OrderBy(k => k).ToList();

        public static readonly List<string> AllManufacturers = new()
        {
            "Analog Devices", "Bourns", "Bourns Inc.", "Diodes", "Diotec",
            "Diotec Semiconductor", "Eaton", "Elmos", "Infineon", "KOA",
            "Kemet", "Kyocera", "Littelfuse", "Lumissil", "MCC", "Microchip",
            "Molex", "Murata", "NXP", "Nexperia", "Nichia", "Onsemi", "Osram",
            "Panasonic", "Panjit International Inc.", "Pulse", "Rohm", "STM",
            "STMicroelectronics", "Samsung", "Stackpole", "Stocko", "TDK",
            "Taiwan Semiconductor", "Taiyo Yuden", "Texas Instruments", "Vishay",
            "Würth", "Yageo", "Yangjie"
        };

        public static readonly List<string> AllFootprints = new()
        {
            "4 PC Pad(ACT45B)", "5023520400-Footprint-1", "CAP-0402", "CAP-0603", "CAP-0805",
            "CAP-1206", "CAP-1210", "D2PAK-3_Ftprnt", "DFN1010D-3 (SOT1215)",
            "DO-214AA(SMB)", "DO-214AB(SMC)", "DO-214AC(SMA)", "DO-219AB(SMF)",
            "ETSSOP-16", "FB-0402", "FB-0603", "FB-0805", "FB-1206", "FB-1206(4-Array)",
            "FCL_Cable_Pad_1X4", "FCL_Terminal_1X4", "HTSSOP-28",
            "IND-12.5x12.5x6(Shielded)", "IND-7.3x6.6x2.8(Shielded)", "IND-7.4x7x4.8(Shielded)",
            "KY DDLN31.FY LED", "LA G6SP.02 LED", "LA.A67F", "LQFP-48",
            "LS E67F_Fprt", "LW.G6SP LED", "LY.G6SP LED", "MFMP 7238 - 202 - W60",
            "NTC-0805", "PG-TSDSO-14", "PowerDI3333-8",
            "RES-0402", "RES-0603", "RES-0805", "RES-1206",
            "Radial, Can", "SC-70", "SMA-FL", "SO-8EP",
            "SOD-123", "SOD-123F", "SOD-123FL", "SOD-123W",
            "SOD-323", "SOD-323EP", "SOD-323F", "SOD-523", "SOD523",
            "SOIC-8", "SOIC16N-EP", "SOT-23", "SOT-26", "SOT-563", "SOT-89",
            "SOT223", "SOT23-3", "SOT403-1", "SOT89", "SOT−363",
            "TFDSO-16", "TQFN-EP(24-SW)", "TSSOP-16",
            "U-DFN2020-6 (Type F)", "WDFN-8", "eTSSOP-16_Fprt"
        };

        public static readonly List<string> AllLibraryRefs = new()
        {
            "AP64200Q_Symbol", "BCR430UW6_Symbol", "CMC_Symbol", "CON Male 2Pin 2.5mm",
            "CapacitorSymbol", "DRV8825PWP_Symbol",
            "Diode_Common_Cathode_Symbol_3pin", "Diode_Schottky_Symbol",
            "Diode_Schottky_Symbol_CathodeLeft", "Diode_Single_Symbol_3pin",
            "Diode_Symbol", "Diode_Symbol_CathodeLeft", "Diode_TVS_Bidirectional_Symbol",
            "Diode_TVS_Unidirectional_Symbol_CathodeLeft",
            "Diode_Zener_Common_Anode_Symbol_3pin", "Diode_Zener_Single_2pin",
            "Diode_Zener_Single_Symbol_3pin", "E522.80_Symbol",
            "FCL_Cable_Pad_1X4", "FCL_Terminal_1X4",
            "FerriteBead(4-Array)Symbol", "FerriteBeadSymbol",
            "IS32LT3144A_Symbol", "IS32LT3152_Symbol", "InductorSymbol",
            "LED_Symbol", "Molex-4Pin", "NCV317BD2TR4G_Symboll",
            "NPN_Transistor_Symbol", "PIC16F15213-E/SN_Symbol",
            "PNP_Transistor_Symbol", "ResistorSymbol",
            "STM32F103C8T6_Symbol", "TLD2142-1EP_Symbol",
            "TLD2382-3ET_Symbol", "TMP112_Symbol"
        };

        // ── Init ──────────────────────────────────────────────────────────
        static ComponentRepository() { }

        public static Task InitializeAsync() => Task.Run(() => Load());

        // ── CRUD ──────────────────────────────────────────────────────────
        public static IReadOnlyList<Component> GetAll() => _components.AsReadOnly();
        public static IReadOnlyList<Component> GetCustom() => _components.Where(c => c.IsCustom).ToList().AsReadOnly();

        public static void Add(Component c)
        {
            c.Id = _nextId++;
            _components.Add(c);

            using var conn = new SqlConnection(ConnectionString);
            conn.Open();

            // Components tablosuna ekle
            using var cmd = new SqlCommand(@"
                INSERT INTO Components
                    (Id, FekaNo, Category, Description, PartNumber, Manufacturer, Footprint,
                     Rating, Process, LibraryRef, FootprintRef, Status, IsCustom, ExtraFields)
                VALUES
                    (@Id, @FekaNo, @Category, @Description, @PartNumber, @Manufacturer, @Footprint,
                     @Rating, @Process, @LibraryRef, @FootprintRef, @Status, @IsCustom, @ExtraFields)", conn);

            BindComponentParams(cmd, c);
            cmd.ExecuteNonQuery();

            // NextId güncelle
            UpdateNextId(conn);
        }

        public static void Update(Component c)
        {
            var idx = _components.FindIndex(x => x.Id == c.Id);
            if (idx >= 0) _components[idx] = c;

            using var conn = new SqlConnection(ConnectionString);
            conn.Open();

            using var cmd = new SqlCommand(@"
                UPDATE Components SET
                    FekaNo = @FekaNo, Category = @Category, Description = @Description,
                    PartNumber = @PartNumber, Manufacturer = @Manufacturer, Footprint = @Footprint,
                    Rating = @Rating, Process = @Process, LibraryRef = @LibraryRef,
                    FootprintRef = @FootprintRef, Status = @Status, IsCustom = @IsCustom,
                    ExtraFields = @ExtraFields
                WHERE Id = @Id", conn);

            BindComponentParams(cmd, c);
            cmd.ExecuteNonQuery();
        }

        public static void Delete(int id)
        {
            _components.RemoveAll(c => c.Id == id);

            using var conn = new SqlConnection(ConnectionString);
            conn.Open();
            using var cmd = new SqlCommand("DELETE FROM Components WHERE Id = @Id", conn);
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.ExecuteNonQuery();
        }

        public static void Approve(int id)
        {
            var c = _components.FirstOrDefault(x => x.Id == id);
            if (c != null)
            {
                c.Status = ComponentStatus.Approved;

                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(
                    "UPDATE Components SET Status = 'Approved' WHERE Id = @Id", conn);
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.ExecuteNonQuery();
            }
        }

        // ── Sorgular (bellek üzerinden — Load() sonrası) ──────────────────
        public static List<Component> Search(string query, string category,
            string rating, string process, string manufacturer)
        {
            var q = query?.ToLower() ?? "";
            return _components.Where(c =>
            {
                if (!string.IsNullOrEmpty(category) && c.Category != category) return false;
                if (!string.IsNullOrEmpty(rating) && c.Rating.ToString() != rating) return false;
                if (!string.IsNullOrEmpty(process) && c.Process.ToString() != process) return false;
                if (!string.IsNullOrEmpty(manufacturer) && c.Manufacturer != manufacturer) return false;
                if (!string.IsNullOrEmpty(q))
                {
                    var text = $"{c.Description} {c.PartNumber} {c.Manufacturer} {c.Category} {c.Footprint}".ToLower();
                    if (!text.Contains(q)) return false;
                }
                return true;
            }).ToList();
        }

        public static Dictionary<string, int> GetCategoryCounts() =>
            _components.GroupBy(c => c.Category).ToDictionary(g => g.Key, g => g.Count());

        public static Dictionary<string, int> GetManufacturerCounts() =>
            _components.Where(c => !string.IsNullOrEmpty(c.Manufacturer))
                       .GroupBy(c => c.Manufacturer)
                       .OrderByDescending(g => g.Count())
                       .Take(10)
                       .ToDictionary(g => g.Key, g => g.Count());

        public static List<string> GetManufacturers() =>
            _components.Select(c => c.Manufacturer)
                       .Where(m => !string.IsNullOrEmpty(m))
                       .Distinct().OrderBy(m => m).ToList();

        public static List<string> GetFootprints() =>
            _components.Select(c => c.Footprint)
                       .Where(f => !string.IsNullOrEmpty(f))
                       .Distinct().OrderBy(f => f).ToList();

        public static List<string> GetLibraryRefs() =>
            _components.Select(c => c.LibraryRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetFootprintRefs() =>
            _components.Select(c => c.FootprintRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetPartNumbers() =>
            _components.Select(c => c.PartNumber)
                       .Where(p => !string.IsNullOrEmpty(p))
                       .Distinct().OrderBy(p => p).ToList();

        public static List<string> GetFootprintsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.Footprint)
                       .Where(f => !string.IsNullOrEmpty(f))
                       .Distinct().OrderBy(f => f).ToList();

        public static List<string> GetManufacturersByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.Manufacturer)
                       .Where(m => !string.IsNullOrEmpty(m))
                       .Distinct().OrderBy(m => m).ToList();

        public static List<string> GetLibraryRefsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.LibraryRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetFootprintRefsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.FootprintRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        // ── Kategori Yönetimi ─────────────────────────────────────────────
        public static void AddCategory(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName)) return;
            if (!CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields[categoryName] = new List<(string Key, string Label)>();
                // Veritabanına alan eklemesi olmadan kategori sadece bellekte tutulur.
                // Alan eklenince CategoryFields tablosuna yazılır (AddFieldToCategory ile).
            }
        }

        public static void DeleteCategory(string categoryName)
        {
            if (CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields.Remove(categoryName);
                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(
                    "DELETE FROM CategoryFields WHERE CategoryName = @Cat", conn);
                cmd.Parameters.AddWithValue("@Cat", categoryName);
                cmd.ExecuteNonQuery();
            }
        }

        public static void AddFieldToCategory(string categoryName, string key, string label)
        {
            if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(key)) return;
            if (!CategoryFields.ContainsKey(categoryName))
                CategoryFields[categoryName] = new();

            if (!CategoryFields[categoryName].Any(f => f.Key == key))
            {
                CategoryFields[categoryName].Add((key, label));
                int order = CategoryFields[categoryName].Count - 1;

                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM CategoryFields WHERE CategoryName=@Cat AND FieldKey=@Key)
                        INSERT INTO CategoryFields (CategoryName, FieldKey, FieldLabel, SortOrder)
                        VALUES (@Cat, @Key, @Label, @Order)", conn);
                cmd.Parameters.AddWithValue("@Cat", categoryName);
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.Parameters.AddWithValue("@Label", label);
                cmd.Parameters.AddWithValue("@Order", order);
                cmd.ExecuteNonQuery();
            }
        }

        public static void RemoveFieldFromCategory(string categoryName, string key)
        {
            if (CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields[categoryName].RemoveAll(f => f.Key == key);

                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(
                    "DELETE FROM CategoryFields WHERE CategoryName=@Cat AND FieldKey=@Key", conn);
                cmd.Parameters.AddWithValue("@Cat", categoryName);
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.ExecuteNonQuery();
            }
        }

        public static void UpdateFieldInCategory(string categoryName, string oldKey, string newKey, string newLabel)
        {
            if (!CategoryFields.ContainsKey(categoryName)) return;
            var fields = CategoryFields[categoryName];
            int idx = fields.FindIndex(f => f.Key == oldKey);
            if (idx >= 0)
            {
                fields[idx] = (newKey, newLabel);

                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    UPDATE CategoryFields SET FieldKey=@NewKey, FieldLabel=@NewLabel
                    WHERE CategoryName=@Cat AND FieldKey=@OldKey", conn);
                cmd.Parameters.AddWithValue("@Cat", categoryName);
                cmd.Parameters.AddWithValue("@OldKey", oldKey);
                cmd.Parameters.AddWithValue("@NewKey", newKey);
                cmd.Parameters.AddWithValue("@NewLabel", newLabel);
                cmd.ExecuteNonQuery();
            }
        }

        // ── CSV Export / Import ───────────────────────────────────────────
        public static string ExportCsv(string categoryFilter = "")
        {
            var recs = string.IsNullOrEmpty(categoryFilter)
                ? _components
                : _components.Where(c => c.Category == categoryFilter).ToList();

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("ID,Category,Description,Part Number,Manufacturer,Footprint,Rating,Process,Library Ref,Footprint Ref,Status");
            foreach (var r in recs)
                sb.AppendLine($"{r.Id},{r.Category},{Esc(r.Description)},{Esc(r.PartNumber)},{Esc(r.Manufacturer)},{r.Footprint},{r.Rating},{r.Process},{r.LibraryRef},{r.FootprintRef},{r.Status}");
            return sb.ToString();
        }

        public static int ImportCsv(string csvContent)
        {
            int count = 0;
            var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return 0;
            var headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
            for (int i = 1; i < lines.Length; i++)
            {
                var vals = lines[i].Split(',');
                if (vals.Length < 2) continue;
                var dict = new Dictionary<string, string>();
                for (int j = 0; j < headers.Length && j < vals.Length; j++)
                    dict[headers[j]] = vals[j].Trim();

                var c = new Component
                {
                    IsCustom     = true,
                    Status       = ComponentStatus.Pending,
                    Category     = dict.GetValueOrDefault("Category", ""),
                    Description  = dict.GetValueOrDefault("Description", ""),
                    PartNumber   = dict.GetValueOrDefault("Part Number", ""),
                    Manufacturer = dict.GetValueOrDefault("Manufacturer", ""),
                    Footprint    = dict.GetValueOrDefault("Footprint", ""),
                };
                if (Enum.TryParse<ComponentRating>(dict.GetValueOrDefault("Rating", "STD"), out var r)) c.Rating = r;
                if (Enum.TryParse<ComponentProcess>(dict.GetValueOrDefault("Process", "SMD"), out var p)) c.Process = p;
                Add(c);
                count++;
            }
            return count;
        }

        // ── Global Fields ─────────────────────────────────────────────────
        public class GlobalFieldDef
        {
            public string Key   { get; set; } = "";
            public string Label { get; set; } = "";
            public string Type  { get; set; } = "text";
        }

        public static List<GlobalFieldDef> GlobalFields { get; private set; } = new();
        public static Dictionary<string, List<string>> GlobalFieldValues { get; private set; } = new();

        public static void AddGlobalField(string key, string label, string type = "text")
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            if (GlobalFields.Any(f => f.Key == key)) return;
            GlobalFields.Add(new GlobalFieldDef { Key = key, Label = label, Type = type });
            if (!GlobalFieldValues.ContainsKey(key))
                GlobalFieldValues[key] = new List<string>();

            using var conn = new SqlConnection(ConnectionString);
            conn.Open();
            using var cmd = new SqlCommand(
                "INSERT INTO GlobalFields (FieldKey, FieldLabel, FieldType) VALUES (@Key, @Label, @Type)", conn);
            cmd.Parameters.AddWithValue("@Key", key);
            cmd.Parameters.AddWithValue("@Label", label);
            cmd.Parameters.AddWithValue("@Type", type);
            cmd.ExecuteNonQuery();
        }

        public static void RemoveGlobalField(string key)
        {
            GlobalFields.RemoveAll(f => f.Key == key);
            GlobalFieldValues.Remove(key);

            using var conn = new SqlConnection(ConnectionString);
            conn.Open();
            using var cmd = new SqlCommand(
                "DELETE FROM GlobalFieldValues WHERE FieldKey=@Key; DELETE FROM GlobalFields WHERE FieldKey=@Key", conn);
            cmd.Parameters.AddWithValue("@Key", key);
            cmd.ExecuteNonQuery();
        }

        public static void RecordGlobalFieldValue(string key, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            if (!GlobalFieldValues.ContainsKey(key))
                GlobalFieldValues[key] = new List<string>();

            if (!GlobalFieldValues[key].Contains(value, StringComparer.OrdinalIgnoreCase))
            {
                GlobalFieldValues[key].Add(value);
                GlobalFieldValues[key].Sort(StringComparer.OrdinalIgnoreCase);

                using var conn = new SqlConnection(ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM GlobalFieldValues WHERE FieldKey=@Key AND FieldValue=@Val)
                        INSERT INTO GlobalFieldValues (FieldKey, FieldValue) VALUES (@Key, @Val)", conn);
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.Parameters.AddWithValue("@Val", value);
                cmd.ExecuteNonQuery();
            }
        }

        // ── Load — tüm veriyi SQL'den belleğe çeker ───────────────────────
        private static void Load()
        {
            try
            {
                using var conn = new SqlConnection(ConnectionString);
                conn.Open();

                // NextId
                using (var cmd = new SqlCommand(
                    "SELECT SettingValue FROM AppSettings WHERE SettingKey='NextId'", conn))
                {
                    var val = cmd.ExecuteScalar();
                    if (val != null) _nextId = int.Parse(val.ToString()!);
                }

                // Components
                _components.Clear();
                using (var cmd = new SqlCommand("SELECT * FROM Components ORDER BY Id", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var extraJson = reader["ExtraFields"] as string;
                        var extra = !string.IsNullOrEmpty(extraJson)
                            ? JsonSerializer.Deserialize<Dictionary<string, string>>(extraJson) ?? new()
                            : new Dictionary<string, string>();

                        _components.Add(new Component
                        {
                            Id           = (int)reader["Id"],
                            FekaNo       = reader["FekaNo"]?.ToString() ?? "",
                            Category     = reader["Category"]?.ToString() ?? "",
                            Description  = reader["Description"]?.ToString() ?? "",
                            PartNumber   = reader["PartNumber"]?.ToString() ?? "",
                            Manufacturer = reader["Manufacturer"]?.ToString() ?? "",
                            Footprint    = reader["Footprint"]?.ToString() ?? "",
                            Rating       = Enum.Parse<ComponentRating>(reader["Rating"]?.ToString() ?? "STD"),
                            Process      = Enum.Parse<ComponentProcess>(reader["Process"]?.ToString() ?? "SMD"),
                            LibraryRef   = reader["LibraryRef"]?.ToString() ?? "",
                            FootprintRef = reader["FootprintRef"]?.ToString() ?? "",
                            Status       = Enum.Parse<ComponentStatus>(reader["Status"]?.ToString() ?? "Pending"),
                            IsCustom     = (bool)reader["IsCustom"],
                            ExtraFields  = extra
                        });
                    }
                }

                // CategoryFields (custom olanlar)
                using (var cmd = new SqlCommand(
                    "SELECT CategoryName, FieldKey, FieldLabel FROM CategoryFields ORDER BY CategoryName, SortOrder", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var catName  = reader["CategoryName"].ToString()!;
                        var fieldKey = reader["FieldKey"].ToString()!;
                        var label    = reader["FieldLabel"].ToString()!;

                        if (!CategoryFields.ContainsKey(catName))
                            CategoryFields[catName] = new();

                        if (!CategoryFields[catName].Any(f => f.Key == fieldKey))
                            CategoryFields[catName].Add((fieldKey, label));
                    }
                }

                // Eksik built-in kategorileri geri ekle
                foreach (var kvp in _builtInCategories)
                {
                    if (!CategoryFields.ContainsKey(kvp.Key))
                        CategoryFields[kvp.Key] = new List<(string, string)>(kvp.Value);
                }

                // GlobalFields
                GlobalFields.Clear();
                using (var cmd = new SqlCommand("SELECT FieldKey, FieldLabel, FieldType FROM GlobalFields", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        GlobalFields.Add(new GlobalFieldDef
                        {
                            Key   = reader["FieldKey"].ToString()!,
                            Label = reader["FieldLabel"].ToString()!,
                            Type  = reader["FieldType"].ToString()!
                        });
                    }
                }

                // GlobalFieldValues
                GlobalFieldValues.Clear();
                using (var cmd = new SqlCommand("SELECT FieldKey, FieldValue FROM GlobalFieldValues ORDER BY FieldKey, FieldValue", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var k = reader["FieldKey"].ToString()!;
                        var v = reader["FieldValue"].ToString()!;
                        if (!GlobalFieldValues.ContainsKey(k))
                            GlobalFieldValues[k] = new();
                        GlobalFieldValues[k].Add(v);
                    }
                }
            }
            catch (Exception ex)
            {
                // Bağlantı hatası — boş bellekle devam et
                System.Windows.Forms.MessageBox.Show(
                    $"Veritabanına bağlanılamadı:\n\n{ex.Message}\n\nLütfen bağlantı ayarlarını kontrol edin.",
                    "Bağlantı Hatası",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                _components = new();
            }
        }

        // ── Yardımcı: NextId güncelle ─────────────────────────────────────
        private static void UpdateNextId(SqlConnection conn)
        {
            using var cmd = new SqlCommand(
                "UPDATE AppSettings SET SettingValue = @Val WHERE SettingKey = 'NextId'", conn);
            cmd.Parameters.AddWithValue("@Val", _nextId.ToString());
            cmd.ExecuteNonQuery();
        }

        // ── Yardımcı: Component parametrelerini bağla ─────────────────────
        private static void BindComponentParams(SqlCommand cmd, Component c)
        {
            var extraJson = c.ExtraFields != null && c.ExtraFields.Count > 0
                ? JsonSerializer.Serialize(c.ExtraFields)
                : null;

            cmd.Parameters.AddWithValue("@Id",           c.Id);
            cmd.Parameters.AddWithValue("@FekaNo",       c.FekaNo);
            cmd.Parameters.AddWithValue("@Category",     c.Category);
            cmd.Parameters.AddWithValue("@Description",  c.Description);
            cmd.Parameters.AddWithValue("@PartNumber",   c.PartNumber);
            cmd.Parameters.AddWithValue("@Manufacturer", c.Manufacturer);
            cmd.Parameters.AddWithValue("@Footprint",    c.Footprint);
            cmd.Parameters.AddWithValue("@Rating",       c.Rating.ToString());
            cmd.Parameters.AddWithValue("@Process",      c.Process.ToString());
            cmd.Parameters.AddWithValue("@LibraryRef",   c.LibraryRef);
            cmd.Parameters.AddWithValue("@FootprintRef", c.FootprintRef);
            cmd.Parameters.AddWithValue("@Status",       c.Status.ToString());
            cmd.Parameters.AddWithValue("@IsCustom",     c.IsCustom);
            cmd.Parameters.AddWithValue("@ExtraFields",  (object?)extraJson ?? DBNull.Value);
        }

        private static string Esc(string s) => s?.Contains(',') == true ? $"\"{s}\"" : s ?? "";
    }
}