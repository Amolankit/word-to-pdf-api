var builder = WebApplication.CreateBuilder(args);

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register custom services
builder.Services.AddScoped<IDocumentService, DocumentService>();
builder.Services.AddScoped<IPdfConversionService, PdfConversionService>();

var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

// Create templates directory if it doesn't exist
var templatesDir = Path.Combine(app.Environment.ContentRootPath, "templates");
if (!Directory.Exists(templatesDir))
{
    Directory.CreateDirectory(templatesDir);
}

app.Run();