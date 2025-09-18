using ClosedXML.Excel;
using Microsoft.AspNetCore.StaticFiles;

var builder = WebApplication.CreateBuilder(args);

// CORS
builder.Services.AddCors(o => o.AddDefaultPolicy(p =>
    p.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod()));

var app = builder.Build();
app.UseCors();

// Static wwwroot (serves index.html)
app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions {
    ContentTypeProvider = new FileExtensionContentTypeProvider()
});

// In-memory store
var store = new DataStore();

app.MapGet("/ping", () => Results.Ok(new { ok = true }));

app.MapPost("/import", async (HttpContext ctx) =>
{
    var payload = await ctx.Request.ReadFromJsonAsync<AllData>() ?? new AllData();
    store.Works = payload.Works ?? new();
    store.Emergencies = payload.Emergencies ?? new();
    store.Grid = payload.Grid ?? new();
    return Results.Ok(new { ok = true, counts = new {
        works = store.Works.Count, emergencies = store.Emergencies.Count, grid = store.Grid.Count
    }});
});

app.MapGet("/export/detail", (string month) =>
{
    var rows = store.Works.Where(w => (w.Date ?? "").StartsWith(month)).ToList();
    var path = Path.Combine(app.Environment.ContentRootPath, "templates", "detail.xlsx");
    using var wb = System.IO.File.Exists(path) ? new ClosedXML.Excel.XLWorkbook(path) : new ClosedXML.Excel.XLWorkbook();
    var ws = wb.Worksheets.Count > 0 ? wb.Worksheet(1) : wb.AddWorksheet("Sheet1");
    ws.RightToLeft = true;

    const int START_ROW = 6;
    int r = START_ROW;
    int idx = 1;
    foreach (var w in rows)
    {
        var spares = (w.Spares?.Count > 0 ? w.Spares : new List<Spare> { new() })!;
        foreach (var sp in spares)
        {
            ws.Cell(r, 1).Value  = idx;
            ws.Cell(r, 2).Value  = w.Weekday;
            ws.Cell(r, 3).Value  = w.Date;
            ws.Cell(r, 4).Value  = w.Region;
            ws.Cell(r, 5).Value  = w.Site;
            ws.Cell(r, 6).Value  = w.SiteOwner;
            ws.Cell(r, 7).Value  = w.JobType;
            ws.Cell(r, 8).Value  = w.Summary;
            ws.Cell(r, 9).Value  = w.OilLiters;
            ws.Cell(r,10).Value  = w.OilFilter ? "✓" : "";
            ws.Cell(r,11).Value  = w.DieselFilter ? "✓" : "";
            ws.Cell(r,12).Value  = w.AirFilter ? "✓" : "";
            ws.Cell(r,13).Value  = w.HoursNow;
            ws.Cell(r,14).Value  = w.HoursDiff;
            ws.Cell(r,15).Value  = w.L1;
            ws.Cell(r,16).Value  = w.L2;
            ws.Cell(r,17).Value  = w.L3;
            ws.Cell(r,18).Value  = w.KwhNow;
            ws.Cell(r,19).Value  = sp?.Name ?? "";
            ws.Cell(r,20).Value  = sp?.Qty  ?? 0;
            ws.Cell(r,21).Value  = w.Executor;
            ws.Cell(r,22).Value  = w.Driver;
            ws.Cell(r,23).Value  = w.Notes;
            r++;
        }
        idx++;
    }
    return StreamWb(wb, $"detail-{month}.xlsx");
});

app.MapGet("/export/summary", (string month) =>
{
    var works = store.Works.Where(w => (w.Date ?? "").StartsWith(month)).ToList();
    var path = Path.Combine(app.Environment.ContentRootPath, "templates", "summary.xlsx");
    using var wb = System.IO.File.Exists(path) ? new ClosedXML.Excel.XLWorkbook(path) : new ClosedXML.Excel.XLWorkbook();
    var ws = wb.Worksheets.Count > 0 ? wb.Worksheet(1) : wb.AddWorksheet("Sheet1");
    ws.RightToLeft = true;

    var tasks = new[] {
        "صيانة مخططة","صيانة دورية","صيانة طارئة","صيانة تفقدية","استلام طوارئ","تعطيل",
        "استلام وتشغيل","ترحيل إنذارات","ربط كهرباء","قراءة عدادات","تكليف عمل","مواد",
        "إصلاحات","أخرى","أخذ القراءات","فحص"
    };
    var regionCols = new Dictionary<string,int> { ["الأمانة"]=4, ["صنعاء"]=5, ["عمران"]=6, ["مأرب"]=7 };
    int START_ROW = 6, COL_TOTAL = 3, r = START_ROW;

    foreach (var t in tasks)
    {
        var subset = works.Where(w => w.JobType == t).ToList();
        ws.Cell(r, COL_TOTAL).Value = subset.Count;
        foreach (var kv in regionCols)
            ws.Cell(r, kv.Value).Value = subset.Count(x => x.Region == kv.Key);
        r++;
    }
    return StreamWb(wb, $"summary-{month}.xlsx");
});

app.MapGet("/export/spares", (string month) =>
{
    var works = store.Works.Where(w => (w.Date ?? "").StartsWith(month)).ToList();
    var path = Path.Combine(app.Environment.ContentRootPath, "templates", "spares.xlsx");
    using var wb = System.IO.File.Exists(path) ? new ClosedXML.Excel.XLWorkbook(path) : new ClosedXML.Excel.XLWorkbook();
    var ws = wb.Worksheets.Count > 0 ? wb.Worksheet(1) : wb.AddWorksheet("Sheet1");
    ws.RightToLeft = true;

    double totalHours = works.Sum(w => w.HoursDiff ?? 0);
    double totalOil   = works.Sum(w => w.OilLiters ?? 0);
    int fOil     = works.Count(w => w.OilFilter);
    int fDiesel  = works.Count(w => w.DieselFilter);
    int fAir     = works.Count(w => w.AirFilter);

    ws.Cell(6,3).Value  = totalHours;
    ws.Cell(7,3).Value  = totalOil;
    ws.Cell(8,3).Value  = fOil;
    ws.Cell(9,3).Value  = fDiesel;
    ws.Cell(10,3).Value = fAir;

    var dict = new Dictionary<string,double>(StringComparer.OrdinalIgnoreCase);
    foreach (var w in works)
        foreach (var sp in w.Spares ?? new())
            if (!string.IsNullOrWhiteSpace(sp.Name))
                dict[sp.Name] = dict.GetValueOrDefault(sp.Name) + (sp.Qty ?? 0);

    int ROW = 20;
    foreach (var kv in dict)
    {
        ws.Cell(ROW,5).Value = kv.Key;
        ws.Cell(ROW,3).Value = kv.Value;
        ROW++;
    }
    return StreamWb(wb, $"spares-{month}.xlsx");
});

app.Run();

static IResult StreamWb(ClosedXML.Excel.XLWorkbook wb, string filename)
{
    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    ms.Position = 0;
    return Results.File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
}

public class AllData
{
    public List<Work>? Works { get; set; }
    public List<Emergency>? Emergencies { get; set; }
    public List<GridData>? Grid { get; set; }
}
public class DataStore
{
    public List<Work> Works { get; set; } = new();
    public List<Emergency> Emergencies { get; set; } = new();
    public List<GridData> Grid { get; set; } = new();
}
public class Spare { public string? Name { get; set; } public double? Qty { get; set; } }
public class Work
{
    public string? Date { get; set; } public string? Weekday { get; set; } public string? Region { get; set; }
    public string? Site { get; set; } public string? SiteOwner { get; set; } public string? JobType { get; set; }
    public string? Summary { get; set; } public double? OilLiters { get; set; } public bool OilFilter { get; set; }
    public bool DieselFilter { get; set; } public bool AirFilter { get; set; } public double? HoursNow { get; set; }
    public double? HoursDiff { get; set; } public double? L1 { get; set; } public double? L2 { get; set; }
    public double? L3 { get; set; } public double? KwhNow { get; set; } public List<Spare>? Spares { get; set; }
    public string? Executor { get; set; } public string? Driver { get; set; } public string? Notes { get; set; }
}
public class Emergency
{
    public string? Date { get; set; } public string? Region { get; set; } public string? Site { get; set; }
    public string? SiteOwner { get; set; } public string? Alarm { get; set; } public string? Source { get; set; }
    public string? Category { get; set; } public string? Notes { get; set; }
}
public class GridData
{
    public string? Date { get; set; } public string? Region { get; set; } public string? Site { get; set; }
    public string? SiteOwner { get; set; } public string? Etype { get; set; } public double? KwhPrev { get; set; }
    public double? KwhNow { get; set; } public double? Kwhr { get; set; } public double? Hours { get; set; }
    public double? KwhDiff { get; set; } public string? SavedAt { get; set; }
}
