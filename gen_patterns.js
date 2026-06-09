const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

// ─── Color palette ──────────────────────────────────────────────────────────
const C = {
  h1bg: "1F3864", h1fg: "FFFFFF",
  h2bg: "2E75B6", h2fg: "FFFFFF",
  h3bg: "D6E4F0", h3fg: "1F3864",
  pyBg:  "EAF4EA", pyBorder: "2E7D32",
  cppBg: "EAF0FB", cppBorder: "1565C0",
  pasBg: "FFF8E1", pasBorder: "F57F17",
  codeFg: "1A1A2E",
  tablHead: "BDD7EE",
  white: "FFFFFF", none: "FFFFFF"
};

const border1 = (color) => ({ style: BorderStyle.SINGLE, size: 4, color });
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

// ─── Helpers ─────────────────────────────────────────────────────────────────
function mono(text, color) {
  return new TextRun({ text, font: "Courier New", size: 18, color: color||C.codeFg });
}
function bold(text, size, color) {
  return new TextRun({ text, bold: true, size: size||24, color: color||"000000" });
}
function run(text, size, color, italic) {
  return new TextRun({ text, size: size||22, color: color||"000000", italics:!!italic });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 40, color: C.h1fg, font:"Arial" })],
    shading: { fill: C.h1bg, type: ShadingType.CLEAR },
    spacing: { before: 360, after: 240 },
    alignment: AlignmentType.CENTER
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 32, color: C.h2fg, font:"Arial" })],
    shading: { fill: C.h2bg, type: ShadingType.CLEAR },
    spacing: { before: 300, after: 200 },
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, size: 26, color: C.h3fg, font:"Arial" })],
    shading: { fill: C.h3bg, type: ShadingType.CLEAR },
    spacing: { before: 240, after: 160 },
  });
}
function para(text, italic) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, italics: !!italic })],
    spacing: { after: 120 }
  });
}

function codeBlock(lines, bgColor, borderColor, label) {
  const b = border1(borderColor);
  const borders = { top:b, bottom:b, left:b, right:b };
  const rows = [];
  // header row
  rows.push(new TableRow({ children: [
    new TableCell({
      borders,
      shading: { fill: borderColor, type: ShadingType.CLEAR },
      margins: { top:60, bottom:60, left:160, right:160 },
      children: [new Paragraph({
        children: [new TextRun({ text: label, bold:true, size:18, color:"FFFFFF", font:"Arial" })]
      })]
    })
  ]}));
  // code rows
  lines.split('\n').forEach(line => {
    rows.push(new TableRow({ children: [
      new TableCell({
        borders,
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top:20, bottom:20, left:200, right:200 },
        children: [new Paragraph({
          children: [mono(line === '' ? ' ' : line)],
          spacing: { before:0, after:0 }
        })]
      })
    ]}));
  });
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows,
    spacing: { before: 160, after: 200 }
  });
}

function pyCode(code) { return codeBlock(code, C.pyBg, C.pyBorder, "Python"); }
function cppCode(code) { return codeBlock(code, C.cppBg, C.cppBorder, "C++"); }
function pasCode(code) { return codeBlock(code, C.pasBg, C.pasBorder, "Turbo Pascal"); }

function space(n) {
  return new Paragraph({ children:[new TextRun(" ")], spacing:{after: n||80} });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── PATTERNS DATA ────────────────────────────────────────────────────────────
const patterns = [

// ══════════════════════════════════════════════════════════════════════════════
// 1. CREATIONAL
// ══════════════════════════════════════════════════════════════════════════════
{
  group: "ПОРОДЖУЮЧІ ПАТЕРНИ (Creational Patterns)",
  name: "1. Singleton (Одинак)",
  intent: "Гарантує, що клас має лише один екземпляр, і надає глобальну точку доступу до нього.",
  when: "Коли потрібен єдиний об'єкт: менеджер конфігурації, логер, пул з'єднань з БД.",
  py: `# Singleton через метаклас
class SingletonMeta(type):
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]

class Logger(metaclass=SingletonMeta):
    def __init__(self):
        self.log = []
    def write(self, msg):
        self.log.append(msg)
        print(f"[LOG] {msg}")

# Тест
a = Logger()
b = Logger()
a.write("Запуск системи")
print(a is b)   # True — той самий об'єкт`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

class Logger {
    Logger() {}
    Logger(const Logger&) = delete;
    Logger& operator=(const Logger&) = delete;
public:
    static Logger& instance() {
        static Logger inst;  // C++11: потокобезпечна ліниво ініціалізація
        return inst;
    }
    void write(const string& msg) {
        cout << "[LOG] " << msg << endl;
    }
};

int main() {
    Logger::instance().write("Запуск системи");
    Logger& a = Logger::instance();
    Logger& b = Logger::instance();
    cout << (&a == &b ? "same" : "diff") << endl; // same
    return 0;
}`,

  pas: `program SingletonDemo;

type
  PLogger = ^TLogger;
  TLogger = object
    class var FInstance: PLogger;
    procedure Write(const Msg: string);
    class function Instance: PLogger;
  end;

var
  Initialized: Boolean = False;
  GlobalLogger: TLogger;

class function TLogger.Instance: PLogger;
begin
  if not Initialized then begin
    GlobalLogger.Write('');  { dummy init }
    FInstance := @GlobalLogger;
    Initialized := True;
  end;
  Result := FInstance;
end;

procedure TLogger.Write(const Msg: string);
begin
  WriteLn('[LOG] ', Msg);
end;

begin
  TLogger.Instance^.Write('Запуск системи');
  { Обидва вказівники — на той самий об'єкт }
  WriteLn(TLogger.Instance = TLogger.Instance);
end.`
},

{
  name: "2. Factory Method (Фабричний метод)",
  intent: "Визначає інтерфейс створення об'єкта, але дозволяє підкласам вирішувати, який клас інстанціювати.",
  when: "Коли заздалегідь невідомо, об'єкти якого класу потрібно створювати.",
  py: `from abc import ABC, abstractmethod

class Transport(ABC):
    @abstractmethod
    def deliver(self): pass

class Truck(Transport):
    def deliver(self): return "Доставка вантажівкою"

class Ship(Transport):
    def deliver(self): return "Доставка кораблем"

class Logistics(ABC):
    @abstractmethod
    def create_transport(self) -> Transport: pass
    def plan_delivery(self):
        t = self.create_transport()
        print(t.deliver())

class RoadLogistics(Logistics):
    def create_transport(self): return Truck()

class SeaLogistics(Logistics):
    def create_transport(self): return Ship()

RoadLogistics().plan_delivery()  # Доставка вантажівкою
SeaLogistics().plan_delivery()   # Доставка кораблем`,

  cpp: `#include <iostream>
#include <memory>
using namespace std;

struct Transport { virtual string deliver() = 0; virtual ~Transport()=default; };
struct Truck : Transport { string deliver() override { return "Доставка вантажівкою"; } };
struct Ship  : Transport { string deliver() override { return "Доставка кораблем";   } };

struct Logistics {
    virtual unique_ptr<Transport> createTransport() = 0;
    void planDelivery() { cout << createTransport()->deliver() << endl; }
    virtual ~Logistics()=default;
};
struct RoadLogistics : Logistics {
    unique_ptr<Transport> createTransport() override { return make_unique<Truck>(); }
};
struct SeaLogistics : Logistics {
    unique_ptr<Transport> createTransport() override { return make_unique<Ship>(); }
};

int main() {
    RoadLogistics().planDelivery();
    SeaLogistics().planDelivery();
}`,

  pas: `program FactoryMethodDemo;

type
  TTransport = class
    function Deliver: string; virtual; abstract;
  end;
  TTruck = class(TTransport)
    function Deliver: string; override;
  end;
  TShip = class(TTransport)
    function Deliver: string; override;
  end;
  TLogistics = class
    function CreateTransport: TTransport; virtual; abstract;
    procedure PlanDelivery;
  end;
  TRoadLogistics = class(TLogistics)
    function CreateTransport: TTransport; override;
  end;

function TTruck.Deliver: string; begin Result := 'Вантажівка'; end;
function TShip.Deliver: string;  begin Result := 'Корабель';  end;

procedure TLogistics.PlanDelivery;
var T: TTransport;
begin
  T := CreateTransport;
  WriteLn(T.Deliver);
  T.Free;
end;

function TRoadLogistics.CreateTransport: TTransport;
begin Result := TTruck.Create; end;

begin
  TRoadLogistics.Create.PlanDelivery;
end.`
},

{
  name: "3. Abstract Factory (Абстрактна фабрика)",
  intent: "Надає інтерфейс для створення родин пов'язаних об'єктів без вказівки їх конкретних класів.",
  when: "Коли система не повинна залежати від того, як створюються, компонуються та представляються продукти.",
  py: `from abc import ABC, abstractmethod

class Button(ABC):
    @abstractmethod
    def render(self): pass

class Checkbox(ABC):
    @abstractmethod
    def render(self): pass

class WinButton(Button):
    def render(self): return "[ Windows Button ]"

class WinCheckbox(Checkbox):
    def render(self): return "[x] Windows Checkbox"

class MacButton(Button):
    def render(self): return "( Mac Button )"

class MacCheckbox(Checkbox):
    def render(self): return "(x) Mac Checkbox"

class GUIFactory(ABC):
    @abstractmethod
    def create_button(self) -> Button: pass
    @abstractmethod
    def create_checkbox(self) -> Checkbox: pass

class WinFactory(GUIFactory):
    def create_button(self):   return WinButton()
    def create_checkbox(self): return WinCheckbox()

class MacFactory(GUIFactory):
    def create_button(self):   return MacButton()
    def create_checkbox(self): return MacCheckbox()

def build_ui(factory: GUIFactory):
    b = factory.create_button()
    c = factory.create_checkbox()
    print(b.render(), c.render())

build_ui(WinFactory())  # [ Windows Button ] [x] Windows Checkbox
build_ui(MacFactory())  # ( Mac Button ) (x) Mac Checkbox`,

  cpp: `#include <iostream>
using namespace std;

struct Button   { virtual string render()=0; virtual ~Button()=default; };
struct Checkbox { virtual string render()=0; virtual ~Checkbox()=default; };

struct WinButton   : Button   { string render() override { return "[ Win Button ]";    } };
struct WinCheckbox : Checkbox { string render() override { return "[x] Win Checkbox";  } };
struct MacButton   : Button   { string render() override { return "( Mac Button )";    } };
struct MacCheckbox : Checkbox { string render() override { return "(x) Mac Checkbox";  } };

struct GUIFactory {
    virtual Button*   createButton()=0;
    virtual Checkbox* createCheckbox()=0;
    virtual ~GUIFactory()=default;
};
struct WinFactory : GUIFactory {
    Button*   createButton()   override { return new WinButton();   }
    Checkbox* createCheckbox() override { return new WinCheckbox(); }
};
struct MacFactory : GUIFactory {
    Button*   createButton()   override { return new MacButton();   }
    Checkbox* createCheckbox() override { return new MacCheckbox(); }
};

int main() {
    GUIFactory* f = new WinFactory();
    cout << f->createButton()->render() << endl;
    cout << f->createCheckbox()->render() << endl;
    delete f;
}`,

  pas: `program AbstractFactoryDemo;

type
  TButton   = class function Render: string; virtual; abstract; end;
  TCheckbox = class function Render: string; virtual; abstract; end;

  TWinButton   = class(TButton)   function Render: string; override; end;
  TWinCheckbox = class(TCheckbox) function Render: string; override; end;
  TMacButton   = class(TButton)   function Render: string; override; end;
  TMacCheckbox = class(TCheckbox) function Render: string; override; end;

  TGUIFactory = class
    function CreateButton:   TButton;   virtual; abstract;
    function CreateCheckbox: TCheckbox; virtual; abstract;
  end;
  TWinFactory = class(TGUIFactory)
    function CreateButton:   TButton;   override;
    function CreateCheckbox: TCheckbox; override;
  end;

function TWinButton.Render:   string; begin Result := '[ Win Button ]';   end;
function TWinCheckbox.Render: string; begin Result := '[x] Win Checkbox'; end;
function TMacButton.Render:   string; begin Result := '( Mac Button )';   end;
function TMacCheckbox.Render: string; begin Result := '(x) Mac Checkbox'; end;

function TWinFactory.CreateButton:   TButton;   begin Result := TWinButton.Create;   end;
function TWinFactory.CreateCheckbox: TCheckbox; begin Result := TWinCheckbox.Create; end;

begin
  with TWinFactory.Create do begin
    WriteLn(CreateButton.Render);
    WriteLn(CreateCheckbox.Render);
    Free;
  end;
end.`
},

{
  name: "4. Builder (Будівник)",
  intent: "Відокремлює конструювання складного об'єкта від його подання, дозволяючи створювати різні подання за одним і тим же процесом.",
  when: "Коли алгоритм створення складного об'єкта не повинен залежати від частин об'єкта та їх збирання.",
  py: `class Pizza:
    def __init__(self):
        self.size = None; self.crust = None
        self.toppings = []
    def __str__(self):
        return f"Pizza({self.size}, {self.crust}, {self.toppings})"

class PizzaBuilder:
    def __init__(self):     self._pizza = Pizza()
    def size(self, s):      self._pizza.size = s;     return self
    def crust(self, c):     self._pizza.crust = c;    return self
    def topping(self, t):   self._pizza.toppings.append(t); return self
    def build(self):        return self._pizza

pizza = (PizzaBuilder()
    .size("large")
    .crust("thin")
    .topping("cheese")
    .topping("mushrooms")
    .build())
print(pizza)  # Pizza(large, thin, ['cheese', 'mushrooms'])`,

  cpp: `#include <iostream>
#include <vector>
#include <string>
using namespace std;

struct Pizza {
    string size, crust;
    vector<string> toppings;
    void print() {
        cout << "Pizza(" << size << ", " << crust << ", [";
        for (auto& t : toppings) cout << t << " ";
        cout << "])" << endl;
    }
};

class PizzaBuilder {
    Pizza p;
public:
    PizzaBuilder& setSize(string s)    { p.size=s;              return *this; }
    PizzaBuilder& setCrust(string c)   { p.crust=c;             return *this; }
    PizzaBuilder& addTopping(string t) { p.toppings.push_back(t); return *this; }
    Pizza build() { return p; }
};

int main() {
    Pizza pizza = PizzaBuilder()
        .setSize("large")
        .setCrust("thin")
        .addTopping("cheese")
        .addTopping("mushrooms")
        .build();
    pizza.print();
}`,

  pas: `program BuilderDemo;

type
  TStringArray = array of string;
  TPizza = record
    Size, Crust: string;
    Toppings: TStringArray;
  end;

  TPizzaBuilder = class
    FPizza: TPizza;
  public
    function SetSize(const S: string): TPizzaBuilder;
    function SetCrust(const C: string): TPizzaBuilder;
    function AddTopping(const T: string): TPizzaBuilder;
    function Build: TPizza;
  end;

function TPizzaBuilder.SetSize(const S: string): TPizzaBuilder;
begin FPizza.Size := S; Result := Self; end;

function TPizzaBuilder.SetCrust(const C: string): TPizzaBuilder;
begin FPizza.Crust := C; Result := Self; end;

function TPizzaBuilder.AddTopping(const T: string): TPizzaBuilder;
var N: Integer;
begin
  N := Length(FPizza.Toppings);
  SetLength(FPizza.Toppings, N+1);
  FPizza.Toppings[N] := T;
  Result := Self;
end;

function TPizzaBuilder.Build: TPizza; begin Result := FPizza; end;

var P: TPizza; I: Integer;
begin
  P := TPizzaBuilder.Create
    .SetSize('large').SetCrust('thin')
    .AddTopping('cheese').AddTopping('mushrooms')
    .Build;
  Write('Pizza(', P.Size, ', ', P.Crust, ', [');
  for I := 0 to High(P.Toppings) do Write(P.Toppings[I],' ');
  WriteLn('])');
end.`
},

{
  name: "5. Prototype (Прототип)",
  intent: "Задає види об'єктів, що створюються, за допомогою екземпляра-прототипу та створює нові об'єкти шляхом копіювання цього прототипу.",
  when: "Коли система повинна бути незалежною від процесу створення об'єктів, а клонування дешевше, ніж створення з нуля.",
  py: `import copy

class Shape:
    def __init__(self, color):
        self.color = color
    def clone(self):
        return copy.deepcopy(self)
    def __str__(self):
        return f"{type(self).__name__}(color={self.color})"

class Circle(Shape):
    def __init__(self, color, radius):
        super().__init__(color)
        self.radius = radius
    def __str__(self):
        return f"Circle(color={self.color}, r={self.radius})"

original = Circle("red", 5)
copy1 = original.clone()
copy1.color = "blue"

print(original)  # Circle(color=red, r=5)
print(copy1)     # Circle(color=blue, r=5)`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

class Shape {
public:
    string color;
    Shape(string c) : color(c) {}
    virtual Shape* clone() const = 0;
    virtual void print() const = 0;
    virtual ~Shape() = default;
};

class Circle : public Shape {
public:
    int radius;
    Circle(string c, int r) : Shape(c), radius(r) {}
    Circle* clone() const override { return new Circle(*this); } // copy ctor
    void print() const override {
        cout << "Circle(color=" << color << ", r=" << radius << ")" << endl;
    }
};

int main() {
    Circle* orig = new Circle("red", 5);
    Circle* copy = orig->clone();
    copy->color = "blue";
    orig->print();   // Circle(color=red, r=5)
    copy->print();   // Circle(color=blue, r=5)
    delete orig; delete copy;
}`,

  pas: `program PrototypeDemo;

type
  TShape = class
    Color: string;
    constructor Create(const C: string);
    function Clone: TShape; virtual; abstract;
    procedure Print; virtual; abstract;
  end;
  TCircle = class(TShape)
    Radius: Integer;
    constructor Create(const C: string; R: Integer);
    function Clone: TShape; override;
    procedure Print; override;
  end;

constructor TShape.Create(const C: string); begin Color := C; end;

constructor TCircle.Create(const C: string; R: Integer);
begin inherited Create(C); Radius := R; end;

function TCircle.Clone: TShape;
begin Result := TCircle.Create(Color, Radius); end;

procedure TCircle.Print;
begin WriteLn('Circle(color=', Color, ', r=', Radius, ')'); end;

var Orig, Copy: TCircle;
begin
  Orig := TCircle.Create('red', 5);
  Copy := TCircle(Orig.Clone);
  Copy.Color := 'blue';
  Orig.Print;   { Circle(color=red, r=5) }
  Copy.Print;   { Circle(color=blue, r=5) }
  Orig.Free; Copy.Free;
end.`
},

// ══════════════════════════════════════════════════════════════════════════════
// 2. STRUCTURAL
// ══════════════════════════════════════════════════════════════════════════════
{
  group: "СТРУКТУРНІ ПАТЕРНИ (Structural Patterns)",
  name: "6. Adapter (Адаптер)",
  intent: "Перетворює інтерфейс класу на інший інтерфейс, якого очікують клієнти.",
  when: "Коли потрібно використовувати існуючий клас, але його інтерфейс не відповідає потрібному.",
  py: `# Є стара система з градусами Фаренгейта
class OldThermometer:
    def get_fahrenheit(self): return 98.6

# Нова система очікує Цельсій
class TemperatureSensor:
    def get_celsius(self): pass

# Адаптер
class ThermometerAdapter(TemperatureSensor):
    def __init__(self, old):
        self._old = old
    def get_celsius(self):
        return (self._old.get_fahrenheit() - 32) * 5 / 9

sensor = ThermometerAdapter(OldThermometer())
print(f"{sensor.get_celsius():.1f}°C")  # 37.0°C`,

  cpp: `#include <iostream>
using namespace std;

class OldThermometer {
public:
    double getFahrenheit() { return 98.6; }
};

class TemperatureSensor {
public:
    virtual double getCelsius() = 0;
    virtual ~TemperatureSensor() = default;
};

class ThermometerAdapter : public TemperatureSensor {
    OldThermometer old;
public:
    double getCelsius() override {
        return (old.getFahrenheit() - 32.0) * 5.0 / 9.0;
    }
};

int main() {
    TemperatureSensor* sensor = new ThermometerAdapter();
    printf("%.1f C\n", sensor->getCelsius());  // 37.0 C
    delete sensor;
}`,

  pas: `program AdapterDemo;

type
  TOldThermometer = class
    function GetFahrenheit: Double;
  end;
  TTemperatureSensor = class
    function GetCelsius: Double; virtual; abstract;
  end;
  TAdapter = class(TTemperatureSensor)
    FOld: TOldThermometer;
    constructor Create;
    function GetCelsius: Double; override;
  end;

function TOldThermometer.GetFahrenheit: Double; begin Result := 98.6; end;

constructor TAdapter.Create; begin FOld := TOldThermometer.Create; end;

function TAdapter.GetCelsius: Double;
begin Result := (FOld.GetFahrenheit - 32) * 5 / 9; end;

var S: TAdapter;
begin
  S := TAdapter.Create;
  WriteLn(S.GetCelsius:4:1, ' C');  { 37.0 C }
  S.Free;
end.`
},

{
  name: "7. Bridge (Міст)",
  intent: "Розділяє абстракцію і реалізацію так, щоб вони могли змінюватися незалежно одна від одної.",
  when: "Коли потрібно уникнути постійного зв'язку між абстракцією і реалізацією, щоб їх можна було змінювати незалежно.",
  py: `from abc import ABC, abstractmethod

class Renderer(ABC):  # Реалізація
    @abstractmethod
    def render_circle(self, r): pass

class VectorRenderer(Renderer):
    def render_circle(self, r): print(f"Вектор: коло r={r}")

class RasterRenderer(Renderer):
    def render_circle(self, r): print(f"Растр: {4*r*r} пікселів")

class Shape(ABC):     # Абстракція
    def __init__(self, renderer: Renderer):
        self.renderer = renderer
    @abstractmethod
    def draw(self): pass

class Circle(Shape):
    def __init__(self, renderer, radius):
        super().__init__(renderer)
        self.radius = radius
    def draw(self):
        self.renderer.render_circle(self.radius)

Circle(VectorRenderer(), 5).draw()  # Вектор: коло r=5
Circle(RasterRenderer(), 5).draw()  # Растр: 100 пікселів`,

  cpp: `#include <iostream>
using namespace std;

struct Renderer { virtual void renderCircle(int r)=0; virtual ~Renderer()=default; };
struct VectorRenderer : Renderer {
    void renderCircle(int r) override { cout<<"Вектор: r="<<r<<endl; }
};
struct RasterRenderer : Renderer {
    void renderCircle(int r) override { cout<<"Растр: "<<4*r*r<<" px"<<endl; }
};

class Shape {
protected: Renderer* rend;
public:
    Shape(Renderer* r): rend(r) {}
    virtual void draw()=0;
    virtual ~Shape()=default;
};
class Circle : public Shape {
    int radius;
public:
    Circle(Renderer* r, int rad): Shape(r), radius(rad) {}
    void draw() override { rend->renderCircle(radius); }
};

int main() {
    VectorRenderer vr; RasterRenderer rr;
    Circle(& vr, 5).draw();
    Circle(& rr, 5).draw();
}`,

  pas: `program BridgeDemo;

type
  TRenderer = class
    procedure RenderCircle(R: Integer); virtual; abstract;
  end;
  TVectorRenderer = class(TRenderer)
    procedure RenderCircle(R: Integer); override;
  end;
  TShape = class
    Renderer: TRenderer;
    constructor Create(R: TRenderer);
    procedure Draw; virtual; abstract;
  end;
  TCircle = class(TShape)
    Radius: Integer;
    constructor Create(R: TRenderer; Rad: Integer);
    procedure Draw; override;
  end;

procedure TVectorRenderer.RenderCircle(R: Integer);
begin WriteLn('Вектор: r=', R); end;

constructor TShape.Create(R: TRenderer); begin Renderer := R; end;
constructor TCircle.Create(R: TRenderer; Rad: Integer);
begin inherited Create(R); Radius := Rad; end;

procedure TCircle.Draw;
begin Renderer.RenderCircle(Radius); end;

var R: TVectorRenderer;
begin
  R := TVectorRenderer.Create;
  TCircle.Create(R, 5).Draw;
  R.Free;
end.`
},

{
  name: "8. Composite (Компоновщик)",
  intent: "Компонує об'єкти у деревоподібні структури для подання ієрархій частина-ціле. Дозволяє однаково працювати з окремими об'єктами і їх композиціями.",
  when: "Коли потрібно представляти ієрархії частина-ціле та клієнтам байдуже, з чим вони працюють — з одним об'єктом чи з їх сукупністю.",
  py: `from abc import ABC, abstractmethod

class Component(ABC):
    @abstractmethod
    def price(self): pass

class Leaf(Component):          # Окремий товар
    def __init__(self, name, p):
        self.name = name; self._price = p
    def price(self): return self._price

class Composite(Component):    # Ящик з товарами
    def __init__(self, name):
        self.name = name; self.children = []
    def add(self, c): self.children.append(c); return self
    def price(self): return sum(c.price() for c in self.children)

box = (Composite("Велика коробка")
    .add(Leaf("Телефон", 10000))
    .add(Composite("Мала коробка")
        .add(Leaf("Зарядка", 200))
        .add(Leaf("Навушники", 500))))

print(f"Всього: {box.price()} грн")  # 10700 грн`,

  cpp: `#include <iostream>
#include <vector>
#include <string>
using namespace std;

struct Component { virtual int price()=0; virtual ~Component()=default; };

struct Leaf : Component {
    string name; int p;
    Leaf(string n, int p): name(n), p(p) {}
    int price() override { return p; }
};

struct Box : Component {
    string name;
    vector<Component*> children;
    Box(string n): name(n) {}
    void add(Component* c) { children.push_back(c); }
    int price() override {
        int total=0;
        for(auto c: children) total += c->price();
        return total;
    }
};

int main() {
    auto* bigBox = new Box("Велика");
    auto* smallBox = new Box("Мала");
    smallBox->add(new Leaf("Зарядка",200));
    smallBox->add(new Leaf("Навушники",500));
    bigBox->add(new Leaf("Телефон",10000));
    bigBox->add(smallBox);
    cout << "Всього: " << bigBox->price() << " грн" << endl;
}`,

  pas: `program CompositeDemo;

type
  TComponent = class
    function Price: Integer; virtual; abstract;
  end;
  TLeaf = class(TComponent)
    FPrice: Integer;
    constructor Create(P: Integer);
    function Price: Integer; override;
  end;
  TBox = class(TComponent)
    Items: array of TComponent;
    procedure Add(C: TComponent);
    function Price: Integer; override;
  end;

constructor TLeaf.Create(P: Integer); begin FPrice := P; end;
function TLeaf.Price: Integer; begin Result := FPrice; end;

procedure TBox.Add(C: TComponent);
var N: Integer;
begin N := Length(Items); SetLength(Items, N+1); Items[N] := C; end;

function TBox.Price: Integer;
var I: Integer;
begin Result := 0; for I := 0 to High(Items) do Inc(Result, Items[I].Price); end;

var Big, Small: TBox;
begin
  Big := TBox.Create; Small := TBox.Create;
  Small.Add(TLeaf.Create(200)); Small.Add(TLeaf.Create(500));
  Big.Add(TLeaf.Create(10000)); Big.Add(Small);
  WriteLn('Всього: ', Big.Price, ' грн');
end.`
},

{
  name: "9. Decorator (Декоратор)",
  intent: "Динамічно додає об'єкту нові обов'язки. Є гнучкою альтернативою успадкуванню при розширенні функціональності.",
  when: "Коли потрібно додавати обов'язки об'єктам динамічно та прозоро, не зачіпаючи інші об'єкти.",
  py: `from abc import ABC, abstractmethod

class Coffee(ABC):
    @abstractmethod
    def cost(self): pass
    @abstractmethod
    def description(self): pass

class SimpleCoffee(Coffee):
    def cost(self): return 20
    def description(self): return "Кава"

class Decorator(Coffee):
    def __init__(self, coffee): self._coffee = coffee
    def cost(self): return self._coffee.cost()
    def description(self): return self._coffee.description()

class Milk(Decorator):
    def cost(self): return super().cost() + 5
    def description(self): return super().description() + " + молоко"

class Sugar(Decorator):
    def cost(self): return super().cost() + 2
    def description(self): return super().description() + " + цукор"

coffee = Sugar(Milk(SimpleCoffee()))
print(coffee.description())  # Кава + молоко + цукор
print(coffee.cost())         # 27`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

struct Coffee {
    virtual int cost()=0;
    virtual string description()=0;
    virtual ~Coffee()=default;
};

struct SimpleCoffee : Coffee {
    int cost() override { return 20; }
    string description() override { return "Кава"; }
};

struct Decorator : Coffee {
    Coffee* wrapped;
    Decorator(Coffee* c): wrapped(c) {}
    int cost() override { return wrapped->cost(); }
    string description() override { return wrapped->description(); }
};

struct Milk : Decorator {
    Milk(Coffee* c): Decorator(c) {}
    int cost() override { return wrapped->cost() + 5; }
    string description() override { return wrapped->description()+" + молоко"; }
};

struct Sugar : Decorator {
    Sugar(Coffee* c): Decorator(c) {}
    int cost() override { return wrapped->cost() + 2; }
    string description() override { return wrapped->description()+" + цукор"; }
};

int main() {
    Coffee* c = new Sugar(new Milk(new SimpleCoffee()));
    cout << c->description() << " = " << c->cost() << " грн" << endl;
}`,

  pas: `program DecoratorDemo;

type
  TCoffee = class
    function Cost: Integer; virtual; abstract;
    function Desc: string; virtual; abstract;
  end;
  TSimpleCoffee = class(TCoffee)
    function Cost: Integer; override;
    function Desc: string;  override;
  end;
  TDecorator = class(TCoffee)
    FWrapped: TCoffee;
    constructor Create(C: TCoffee);
    function Cost: Integer; override;
    function Desc: string;  override;
  end;
  TMilk = class(TDecorator)
    function Cost: Integer; override;
    function Desc: string;  override;
  end;

function TSimpleCoffee.Cost: Integer; begin Result := 20; end;
function TSimpleCoffee.Desc: string;  begin Result := 'Кава'; end;

constructor TDecorator.Create(C: TCoffee); begin FWrapped := C; end;
function TDecorator.Cost: Integer; begin Result := FWrapped.Cost; end;
function TDecorator.Desc: string;  begin Result := FWrapped.Desc; end;

function TMilk.Cost: Integer; begin Result := FWrapped.Cost + 5; end;
function TMilk.Desc: string;  begin Result := FWrapped.Desc + ' + молоко'; end;

var C: TMilk;
begin
  C := TMilk.Create(TSimpleCoffee.Create);
  WriteLn(C.Desc, ' = ', C.Cost, ' грн');
  C.Free;
end.`
},

{
  name: "10. Facade (Фасад)",
  intent: "Надає уніфікований інтерфейс до набору інтерфейсів підсистеми. Визначає інтерфейс вищого рівня, що спрощує використання підсистеми.",
  when: "Коли потрібно надати простий інтерфейс до складної підсистеми.",
  py: `class CPU:
    def freeze(self): print("CPU: заморожений")
    def execute(self): print("CPU: виконання")

class RAM:
    def load(self, data): print(f"RAM: завантажено '{data}'")

class HDD:
    def read(self): return "ОС дані"

class ComputerFacade:
    def __init__(self):
        self.cpu = CPU(); self.ram = RAM(); self.hdd = HDD()
    def start(self):
        self.cpu.freeze()
        self.ram.load(self.hdd.read())
        self.cpu.execute()

ComputerFacade().start()`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

class CPU { public: void freeze(){cout<<"CPU freeze\n";} void execute(){cout<<"CPU run\n";} };
class RAM { public: void load(string d){cout<<"RAM: "<<d<<"\n";} };
class HDD { public: string read(){ return "OS data"; } };

class Computer {
    CPU cpu; RAM ram; HDD hdd;
public:
    void start() {
        cpu.freeze();
        ram.load(hdd.read());
        cpu.execute();
    }
};

int main() { Computer().start(); }`,

  pas: `program FacadeDemo;

type
  TCPU = class procedure Freeze; procedure Execute; end;
  TRAM = class procedure Load(const D: string); end;
  THDD = class function Read: string; end;
  TComputer = class
    CPU: TCPU; RAM: TRAM; HDD: THDD;
    constructor Create;
    procedure Start;
  end;

procedure TCPU.Freeze;  begin WriteLn('CPU: freeze'); end;
procedure TCPU.Execute; begin WriteLn('CPU: run');    end;
procedure TRAM.Load(const D: string); begin WriteLn('RAM: ', D); end;
function  THDD.Read: string; begin Result := 'OS data'; end;

constructor TComputer.Create;
begin CPU:=TCPU.Create; RAM:=TRAM.Create; HDD:=THDD.Create; end;

procedure TComputer.Start;
begin CPU.Freeze; RAM.Load(HDD.Read); CPU.Execute; end;

begin TComputer.Create.Start; end.`
},

{
  name: "11. Flyweight (Легковаговик)",
  intent: "Використовує спільний стан для ефективної підтримки великої кількості дрібних об'єктів.",
  when: "Коли використовується велика кількість схожих об'єктів, і це спричиняє неприйнятні витрати пам'яті.",
  py: `class TreeType:  # Flyweight — спільний стан
    def __init__(self, name, color):
        self.name = name; self.color = color
    def draw(self, x, y):
        print(f"Дерево {self.name}({self.color}) @ ({x},{y})")

class TreeFactory:
    _types = {}
    @classmethod
    def get(cls, name, color):
        key = (name, color)
        if key not in cls._types:
            cls._types[key] = TreeType(name, color)
        return cls._types[key]

class Tree:    # Унікальний стан (x, y)
    def __init__(self, x, y, tree_type):
        self.x=x; self.y=y; self.type=tree_type
    def draw(self): self.type.draw(self.x, self.y)

forest = [
    Tree(1, 2, TreeFactory.get("Дуб", "зелений")),
    Tree(3, 4, TreeFactory.get("Дуб", "зелений")),  # той самий Flyweight
    Tree(5, 6, TreeFactory.get("Сосна", "темний")),
]
for t in forest: t.draw()
print(f"Унікальних типів: {len(TreeFactory._types)}")  # 2`,

  cpp: `#include <iostream>
#include <map>
#include <string>
using namespace std;

class TreeType {
    string name, color;
public:
    TreeType(string n, string c): name(n), color(c) {}
    void draw(int x, int y) {
        cout<<"Дерево "<<name<<"("<<color<<") @ ("<<x<<","<<y<<")\n";
    }
};

class TreeFactory {
    static map<pair<string,string>, TreeType*> types;
public:
    static TreeType* get(string n, string c) {
        auto key = make_pair(n,c);
        if(!types.count(key)) types[key] = new TreeType(n,c);
        return types[key];
    }
    static int count() { return types.size(); }
};
map<pair<string,string>, TreeType*> TreeFactory::types;

int main() {
    TreeFactory::get("Дуб","зелений")->draw(1,2);
    TreeFactory::get("Дуб","зелений")->draw(3,4);
    TreeFactory::get("Сосна","темний")->draw(5,6);
    cout << "Типів: " << TreeFactory::count() << endl;
}`,

  pas: `program FlyweightDemo;

type
  TTreeType = class
    Name, Color: string;
    constructor Create(N, C: string);
    procedure Draw(X, Y: Integer);
  end;

constructor TTreeType.Create(N, C: string); begin Name:=N; Color:=C; end;
procedure TTreeType.Draw(X, Y: Integer);
begin WriteLn('Дерево ',Name,'(',Color,') @ (',X,',',Y,')'); end;

{ Спрощена фабрика — два відомих типи }
var Oak, Pine: TTreeType;
begin
  Oak  := TTreeType.Create('Дуб', 'зелений');
  Pine := TTreeType.Create('Сосна', 'темний');
  Oak.Draw(1,2);
  Oak.Draw(3,4);   { той самий об'єкт }
  Pine.Draw(5,6);
  WriteLn('Типів: 2');
  Oak.Free; Pine.Free;
end.`
},

{
  name: "12. Proxy (Замісник)",
  intent: "Надає об'єкту-заміснику або об'єкту-посереднику для управління доступом до іншого об'єкта.",
  when: "Коли потрібне більш гнучке або інтелектуальне посилання на об'єкт, ніж звичайний покажчик.",
  py: `from abc import ABC, abstractmethod

class Image(ABC):
    @abstractmethod
    def display(self): pass

class RealImage(Image):
    def __init__(self, file):
        self.file = file
        print(f"Завантаження {file}...")  # Дорога операція
    def display(self): print(f"Відображення {self.file}")

class ProxyImage(Image):         # Ліниве завантаження
    def __init__(self, file):
        self.file = file; self._real = None
    def display(self):
        if not self._real:
            self._real = RealImage(self.file)
        self._real.display()

img = ProxyImage("photo.jpg")
print("Зображення створено")
img.display()   # Тут відбувається завантаження
img.display()   # Повторне — без завантаження`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

struct Image { virtual void display()=0; virtual ~Image()=default; };

struct RealImage : Image {
    string file;
    RealImage(string f): file(f) { cout<<"Завантаження "<<f<<"\n"; }
    void display() override { cout<<"Відображення "<<file<<"\n"; }
};

struct ProxyImage : Image {
    string file; RealImage* real = nullptr;
    ProxyImage(string f): file(f) {}
    void display() override {
        if(!real) real = new RealImage(file);
        real->display();
    }
    ~ProxyImage() { delete real; }
};

int main() {
    Image* img = new ProxyImage("photo.jpg");
    cout << "Об'єкт створено\n";
    img->display();  // завантаження тут
    img->display();  // без завантаження
    delete img;
}`,

  pas: `program ProxyDemo;

type
  TImage = class
    procedure Display; virtual; abstract;
  end;
  TRealImage = class(TImage)
    FileName: string;
    constructor Create(const F: string);
    procedure Display; override;
  end;
  TProxyImage = class(TImage)
    FileName: string;
    FReal: TRealImage;
    constructor Create(const F: string);
    procedure Display; override;
  end;

constructor TRealImage.Create(const F: string);
begin FileName := F; WriteLn('Завантаження ', F); end;
procedure TRealImage.Display;
begin WriteLn('Відображення ', FileName); end;

constructor TProxyImage.Create(const F: string);
begin FileName := F; FReal := nil; end;

procedure TProxyImage.Display;
begin
  if FReal = nil then FReal := TRealImage.Create(FileName);
  FReal.Display;
end;

var Img: TProxyImage;
begin
  Img := TProxyImage.Create('photo.jpg');
  WriteLn('Об''єкт створено');
  Img.Display;
  Img.Display;
  Img.Free;
end.`
},

// ══════════════════════════════════════════════════════════════════════════════
// 3. BEHAVIORAL
// ══════════════════════════════════════════════════════════════════════════════
{
  group: "ПОВЕДІНКОВІ ПАТЕРНИ (Behavioral Patterns)",
  name: "13. Chain of Responsibility (Ланцюжок відповідальності)",
  intent: "Дає можливість передавати запити по ланцюжку обробників, де кожен вирішує — обробляти чи передавати далі.",
  when: "Коли є більше одного об'єкта, здатного обробити запит, і конкретний обробник не відомий заздалегідь.",
  py: `from abc import ABC, abstractmethod

class Handler(ABC):
    def __init__(self): self._next = None
    def set_next(self, handler):
        self._next = handler; return handler
    def handle(self, request):
        if self._next: return self._next.handle(request)

class LowHandler(Handler):
    def handle(self, lvl):
        if lvl <= 1: print(f"LowHandler обробив рівень {lvl}")
        else: super().handle(lvl)

class MidHandler(Handler):
    def handle(self, lvl):
        if lvl <= 3: print(f"MidHandler обробив рівень {lvl}")
        else: super().handle(lvl)

class HighHandler(Handler):
    def handle(self, lvl): print(f"HighHandler обробив рівень {lvl}")

low = LowHandler()
low.set_next(MidHandler()).set_next(HighHandler())

for level in [1, 2, 5]:
    low.handle(level)`,

  cpp: `#include <iostream>
using namespace std;

class Handler {
protected: Handler* next = nullptr;
public:
    Handler* setNext(Handler* h) { next=h; return h; }
    virtual void handle(int lvl) {
        if(next) next->handle(lvl);
    }
};
class Low  : public Handler { public: void handle(int lvl) override {
    if(lvl<=1) cout<<"Low: "<<lvl<<"\n"; else Handler::handle(lvl); }};
class Mid  : public Handler { public: void handle(int lvl) override {
    if(lvl<=3) cout<<"Mid: "<<lvl<<"\n"; else Handler::handle(lvl); }};
class High : public Handler { public: void handle(int lvl) override {
    cout<<"High: "<<lvl<<"\n"; }};

int main() {
    Low l; Mid m; High h;
    l.setNext(&m)->setNext(&h);
    for(int lv: {1,2,5}) l.handle(lv);
}`,

  pas: `program ChainDemo;

type
  THandler = class
    Next: THandler;
    procedure Handle(Lvl: Integer); virtual;
  end;
  TLow  = class(THandler) procedure Handle(Lvl: Integer); override; end;
  TMid  = class(THandler) procedure Handle(Lvl: Integer); override; end;
  THigh = class(THandler) procedure Handle(Lvl: Integer); override; end;

procedure THandler.Handle(Lvl: Integer);
begin if Assigned(Next) then Next.Handle(Lvl); end;

procedure TLow.Handle(Lvl: Integer);
begin if Lvl<=1 then WriteLn('Low: ',Lvl) else inherited; end;
procedure TMid.Handle(Lvl: Integer);
begin if Lvl<=3 then WriteLn('Mid: ',Lvl) else inherited; end;
procedure THigh.Handle(Lvl: Integer);
begin WriteLn('High: ',Lvl); end;

var L: TLow; M: TMid; H: THigh;
begin
  L:=TLow.Create; M:=TMid.Create; H:=THigh.Create;
  L.Next:=M; M.Next:=H;
  L.Handle(1); L.Handle(2); L.Handle(5);
  L.Free; M.Free; H.Free;
end.`
},

{
  name: "14. Command (Команда)",
  intent: "Інкапсулює запит як об'єкт, дозволяючи параметризувати клієнтів з різними запитами, чергами, логуванням та скасуванням операцій.",
  when: "Коли потрібно параметризувати об'єкти виконуваною дією або підтримати скасування операцій.",
  py: `from abc import ABC, abstractmethod

class Command(ABC):
    @abstractmethod
    def execute(self): pass
    @abstractmethod
    def undo(self): pass

class TextEditor:
    def __init__(self): self.text = ""
    def write(self, s): self.text += s
    def delete(self, n): self.text = self.text[:-n]

class WriteCommand(Command):
    def __init__(self, editor, text):
        self.editor=editor; self.text=text
    def execute(self): self.editor.write(self.text)
    def undo(self):    self.editor.delete(len(self.text))

editor = TextEditor()
history = []

cmd = WriteCommand(editor, "Привіт")
cmd.execute(); history.append(cmd)
cmd2 = WriteCommand(editor, " світ")
cmd2.execute(); history.append(cmd2)

print(editor.text)      # Привіт світ
history.pop().undo()
print(editor.text)      # Привіт`,

  cpp: `#include <iostream>
#include <string>
#include <vector>
using namespace std;

struct Command { virtual void execute()=0; virtual void undo()=0; virtual ~Command()=default; };

class Editor { public: string text; };

class WriteCmd : public Command {
    Editor& e; string txt;
public:
    WriteCmd(Editor& ed, string t): e(ed), txt(t) {}
    void execute() override { e.text += txt; }
    void undo()    override { e.text.resize(e.text.size()-txt.size()); }
};

int main() {
    Editor ed;
    vector<Command*> history;
    auto* c1 = new WriteCmd(ed, "Привіт");
    c1->execute(); history.push_back(c1);
    auto* c2 = new WriteCmd(ed, " світ");
    c2->execute(); history.push_back(c2);

    cout << ed.text << endl;       // Привіт світ
    history.back()->undo();
    cout << ed.text << endl;       // Привіт
}`,

  pas: `program CommandDemo;

type
  TEditor = class Text: string; end;
  TCommand = class
    procedure Execute; virtual; abstract;
    procedure Undo; virtual; abstract;
  end;
  TWriteCmd = class(TCommand)
    E: TEditor; Txt: string;
    constructor Create(Ed: TEditor; T: string);
    procedure Execute; override;
    procedure Undo; override;
  end;

constructor TWriteCmd.Create(Ed: TEditor; T: string);
begin E := Ed; Txt := T; end;

procedure TWriteCmd.Execute; begin E.Text := E.Text + Txt; end;
procedure TWriteCmd.Undo;
begin Delete(E.Text, Length(E.Text)-Length(Txt)+1, Length(Txt)); end;

var Ed: TEditor; C1, C2: TWriteCmd;
begin
  Ed := TEditor.Create;
  C1 := TWriteCmd.Create(Ed, 'Привіт'); C1.Execute;
  C2 := TWriteCmd.Create(Ed, ' світ'); C2.Execute;
  WriteLn(Ed.Text);
  C2.Undo;
  WriteLn(Ed.Text);
  Ed.Free; C1.Free; C2.Free;
end.`
},

{
  name: "15. Iterator (Ітератор)",
  intent: "Надає спосіб послідовного доступу до елементів агрегатного об'єкта без розкриття його внутрішньої структури.",
  when: "Коли потрібен уніфікований спосіб перебору різних колекцій.",
  py: `class WordIterator:
    def __init__(self, text):
        self._words = text.split()
        self._index = 0
    def __iter__(self): return self
    def __next__(self):
        if self._index >= len(self._words):
            raise StopIteration
        word = self._words[self._index]
        self._index += 1
        return word

for word in WordIterator("Патерни проєктування це круто"):
    print(word)`,

  cpp: `#include <iostream>
#include <vector>
#include <string>
using namespace std;

class WordIterator {
    vector<string> words;
    size_t idx = 0;
public:
    WordIterator(vector<string> w): words(w) {}
    bool hasNext() { return idx < words.size(); }
    string next()  { return words[idx++]; }
};

int main() {
    WordIterator it({"Патерни","проєктування","це","круто"});
    while(it.hasNext())
        cout << it.next() << " ";
    cout << endl;
}`,

  pas: `program IteratorDemo;

type
  TWords = array of string;
  TWordIterator = class
    FWords: TWords;
    FIdx: Integer;
    constructor Create(W: TWords);
    function HasNext: Boolean;
    function Next: string;
  end;

constructor TWordIterator.Create(W: TWords);
begin FWords := W; FIdx := 0; end;

function TWordIterator.HasNext: Boolean;
begin Result := FIdx < Length(FWords); end;

function TWordIterator.Next: string;
begin Result := FWords[FIdx]; Inc(FIdx); end;

var It: TWordIterator; W: TWords;
begin
  SetLength(W, 3); W[0]:='Патерни'; W[1]:='це'; W[2]:='круто';
  It := TWordIterator.Create(W);
  while It.HasNext do Write(It.Next, ' ');
  WriteLn;
  It.Free;
end.`
},

{
  name: "16. Mediator (Посередник)",
  intent: "Визначає об'єкт, що інкапсулює взаємодію між об'єктами. Об'єкти більше не посилаються один на одного явно.",
  when: "Коли складна взаємодія об'єктів призводить до заплутаних взаємозалежностей.",
  py: `class ChatMediator:
    def __init__(self): self._users = []
    def add(self, user): self._users.append(user)
    def send(self, msg, sender):
        for u in self._users:
            if u is not sender:
                u.receive(msg, sender.name)

class User:
    def __init__(self, name, chat):
        self.name = name; self._chat = chat
        chat.add(self)
    def send(self, msg): self._chat.send(msg, self)
    def receive(self, msg, frm): print(f"{self.name} отримав від {frm}: {msg}")

chat = ChatMediator()
alice = User("Аліса", chat)
bob   = User("Боб",   chat)
carol = User("Кароль",chat)

alice.send("Привіт всім!")`,

  cpp: `#include <iostream>
#include <vector>
#include <string>
using namespace std;

class User;
class ChatMediator {
    vector<User*> users;
public:
    void addUser(User* u) { users.push_back(u); }
    void send(string msg, User* sender);
};

class User {
    string name; ChatMediator& chat;
public:
    User(string n, ChatMediator& c): name(n), chat(c) { c.addUser(this); }
    void send(string msg) { chat.send(msg, this); }
    void receive(string msg, string from) {
        cout << name << " отримав від " << from << ": " << msg << "\n";
    }
    string getName() { return name; }
};

void ChatMediator::send(string msg, User* sender) {
    for(auto* u: users) if(u != sender) u->receive(msg, sender->getName());
}

int main() {
    ChatMediator chat;
    User alice("Аліса",chat), bob("Боб",chat);
    alice.send("Привіт!");
}`,

  pas: `program MediatorDemo;

type
  TUser = class; { forward }
  TChat = class
    Users: array of TUser;
    procedure Add(U: TUser);
    procedure Send(const Msg: string; Sender: TUser);
  end;
  TUser = class
    Name: string; Chat: TChat;
    constructor Create(const N: string; C: TChat);
    procedure Send(const Msg: string);
    procedure Receive(const Msg, From: string);
  end;

procedure TChat.Add(U: TUser);
var N: Integer;
begin N:=Length(Users); SetLength(Users,N+1); Users[N]:=U; end;

procedure TChat.Send(const Msg: string; Sender: TUser);
var I: Integer;
begin
  for I:=0 to High(Users) do
    if Users[I] <> Sender then Users[I].Receive(Msg, Sender.Name);
end;

constructor TUser.Create(const N: string; C: TChat);
begin Name:=N; Chat:=C; C.Add(Self); end;

procedure TUser.Send(const Msg: string); begin Chat.Send(Msg, Self); end;
procedure TUser.Receive(const Msg, From: string);
begin WriteLn(Name, ' отримав від ', From, ': ', Msg); end;

var C: TChat; Alice, Bob: TUser;
begin
  C:=TChat.Create;
  Alice:=TUser.Create('Аліса',C);
  Bob:=TUser.Create('Боб',C);
  Alice.Send('Привіт!');
  C.Free;
end.`
},

{
  name: "17. Memento (Знімок)",
  intent: "Не порушуючи інкапсуляцію, фіксує та виносить за межі об'єкта його внутрішній стан так, щоб пізніше можна було відновити цей стан.",
  when: "Коли потрібно реалізувати скасування операцій (undo).",
  py: `class Memento:
    def __init__(self, state): self._state = state
    def get_state(self): return self._state

class TextEditor:
    def __init__(self): self._text = ""
    def write(self, s): self._text += s
    def save(self): return Memento(self._text)
    def restore(self, m): self._text = m.get_state()
    def __str__(self): return self._text

editor = TextEditor()
editor.write("Версія 1")
snap1 = editor.save()

editor.write(" -> Версія 2")
snap2 = editor.save()

editor.write(" -> ПОМИЛКА")
print(editor)               # Версія 1 -> Версія 2 -> ПОМИЛКА

editor.restore(snap2)
print(editor)               # Версія 1 -> Версія 2`,

  cpp: `#include <iostream>
#include <string>
using namespace std;

class Memento {
    string state;
public:
    Memento(string s): state(s) {}
    string getState() { return state; }
};

class Editor {
    string text;
public:
    void write(string s) { text += s; }
    Memento save() { return Memento(text); }
    void restore(Memento m) { text = m.getState(); }
    void print() { cout << text << endl; }
};

int main() {
    Editor ed;
    ed.write("Версія 1");
    Memento s1 = ed.save();
    ed.write(" -> Версія 2");
    Memento s2 = ed.save();
    ed.write(" -> ПОМИЛКА");
    ed.print();          // Версія 1 -> Версія 2 -> ПОМИЛКА
    ed.restore(s2);
    ed.print();          // Версія 1 -> Версія 2
}`,

  pas: `program MementoDemo;

type
  TMemento = class State: string; end;
  TEditor = class
    Text: string;
    function Save: TMemento;
    procedure Restore(M: TMemento);
    procedure Write(const S: string);
  end;

function TEditor.Save: TMemento;
begin Result:=TMemento.Create; Result.State:=Text; end;
procedure TEditor.Restore(M: TMemento); begin Text:=M.State; end;
procedure TEditor.Write(const S: string); begin Text:=Text+S; end;

var Ed: TEditor; S1, S2: TMemento;
begin
  Ed:=TEditor.Create;
  Ed.Write('Версія 1'); S1:=Ed.Save;
  Ed.Write(' -> Версія 2'); S2:=Ed.Save;
  Ed.Write(' -> ПОМИЛКА');
  WriteLn(Ed.Text);
  Ed.Restore(S2);
  WriteLn(Ed.Text);
  Ed.Free; S1.Free; S2.Free;
end.`
},

{
  name: "18. Observer (Спостерігач)",
  intent: "Визначає залежність один-до-багатьох між об'єктами, де при зміні стану одного об'єкта всі залежні від нього оповіщуються автоматично.",
  when: "Коли зміна стану одного об'єкта вимагає оновлення інших, і кількість цих об'єктів заздалегідь невідома.",
  py: `from abc import ABC, abstractmethod

class Observer(ABC):
    @abstractmethod
    def update(self, temp): pass

class Subject:
    def __init__(self):
        self._observers = []; self._temp = 0
    def attach(self, o): self._observers.append(o)
    def set_temp(self, t):
        self._temp = t; self._notify()
    def _notify(self):
        for o in self._observers: o.update(self._temp)

class PhoneDisplay(Observer):
    def update(self, t): print(f"Телефон: {t}°C")

class WindowDisplay(Observer):
    def update(self, t): print(f"Вікно: {t}°C")

station = Subject()
station.attach(PhoneDisplay())
station.attach(WindowDisplay())

station.set_temp(25)   # Обидва дисплеї оновляться`,

  cpp: `#include <iostream>
#include <vector>
using namespace std;

struct Observer { virtual void update(int t)=0; virtual ~Observer()=default; };

class Subject {
    vector<Observer*> obs; int temp=0;
public:
    void attach(Observer* o) { obs.push_back(o); }
    void setTemp(int t) { temp=t; for(auto* o:obs) o->update(t); }
};

struct Phone  : Observer { void update(int t) override { cout<<"Телефон: "<<t<<"°C\n"; } };
struct Window : Observer { void update(int t) override { cout<<"Вікно: "<<t<<"°C\n";   } };

int main() {
    Subject station;
    Phone ph; Window wn;
    station.attach(&ph); station.attach(&wn);
    station.setTemp(25);
}`,

  pas: `program ObserverDemo;

type
  TObserver = class procedure Update(T: Integer); virtual; abstract; end;
  TSubject = class
    Obs: array of TObserver;
    Temp: Integer;
    procedure Attach(O: TObserver);
    procedure SetTemp(T: Integer);
    procedure Notify;
  end;
  TPhone  = class(TObserver) procedure Update(T: Integer); override; end;
  TWindow = class(TObserver) procedure Update(T: Integer); override; end;

procedure TSubject.Attach(O: TObserver);
var N: Integer;
begin N:=Length(Obs); SetLength(Obs,N+1); Obs[N]:=O; end;

procedure TSubject.SetTemp(T: Integer); begin Temp:=T; Notify; end;

procedure TSubject.Notify;
var I: Integer;
begin for I:=0 to High(Obs) do Obs[I].Update(Temp); end;

procedure TPhone.Update(T: Integer);  begin WriteLn('Телефон: ',T,'°C'); end;
procedure TWindow.Update(T: Integer); begin WriteLn('Вікно: ',T,'°C');   end;

var S: TSubject; P: TPhone; W: TWindow;
begin
  S:=TSubject.Create; P:=TPhone.Create; W:=TWindow.Create;
  S.Attach(P); S.Attach(W);
  S.SetTemp(25);
  S.Free; P.Free; W.Free;
end.`
},

{
  name: "19. State (Стан)",
  intent: "Дозволяє об'єкту змінювати свою поведінку при зміні внутрішнього стану — здається, що об'єкт змінив свій клас.",
  when: "Коли поведінка об'єкта залежить від його стану і повинна змінюватися під час виконання.",
  py: `from abc import ABC, abstractmethod

class State(ABC):
    @abstractmethod
    def handle(self, context): pass

class GreenLight(State):
    def handle(self, ctx):
        print("Зелений — їдемо!")
        ctx.state = YellowLight()

class YellowLight(State):
    def handle(self, ctx):
        print("Жовтий — обережно!")
        ctx.state = RedLight()

class RedLight(State):
    def handle(self, ctx):
        print("Червоний — стоїмо!")
        ctx.state = GreenLight()

class TrafficLight:
    def __init__(self): self.state = GreenLight()
    def change(self):   self.state.handle(self)

light = TrafficLight()
for _ in range(4): light.change()`,

  cpp: `#include <iostream>
using namespace std;

class Context; // forward
struct State { virtual void handle(Context& ctx)=0; virtual ~State()=default; };

class GreenLight; class YellowLight; class RedLight;

class Context {
public: State* state;
    Context();
    void change() { state->handle(*this); }
    ~Context() { delete state; }
};

struct GreenLight : State {
    void handle(Context& ctx) override;
};
struct YellowLight : State {
    void handle(Context& ctx) override {
        cout<<"Жовтий!\n"; delete ctx.state; ctx.state=new RedLight();
    }
};
struct RedLight : State {
    void handle(Context& ctx) override {
        cout<<"Червоний!\n"; delete ctx.state; ctx.state=new GreenLight();
    }
};

void GreenLight::handle(Context& ctx) {
    cout<<"Зелений!\n"; delete ctx.state; ctx.state=new YellowLight();
}
Context::Context(): state(new GreenLight()) {}

int main() {
    Context light;
    for(int i=0;i<4;i++) light.change();
}`,

  pas: `program StateDemo;

type
  TContext = class; { forward }
  TState = class
    procedure Handle(Ctx: TContext); virtual; abstract;
  end;
  TContext = class
    State: TState;
    constructor Create;
    procedure Change;
  end;
  TGreenLight  = class(TState) procedure Handle(Ctx: TContext); override; end;
  TYellowLight = class(TState) procedure Handle(Ctx: TContext); override; end;
  TRedLight    = class(TState) procedure Handle(Ctx: TContext); override; end;

constructor TContext.Create; begin State := TGreenLight.Create; end;
procedure TContext.Change; begin State.Handle(Self); end;

procedure TGreenLight.Handle(Ctx: TContext);
begin WriteLn('Зелений!'); Ctx.State.Free; Ctx.State:=TYellowLight.Create; end;
procedure TYellowLight.Handle(Ctx: TContext);
begin WriteLn('Жовтий!');  Ctx.State.Free; Ctx.State:=TRedLight.Create;    end;
procedure TRedLight.Handle(Ctx: TContext);
begin WriteLn('Червоний!');Ctx.State.Free; Ctx.State:=TGreenLight.Create;  end;

var Light: TContext; I: Integer;
begin
  Light:=TContext.Create;
  for I:=1 to 4 do Light.Change;
  Light.Free;
end.`
},

{
  name: "20. Strategy (Стратегія)",
  intent: "Визначає сімейство алгоритмів, інкапсулює кожен з них та робить їх взаємозамінними, дозволяючи змінювати алгоритм незалежно від клієнтів.",
  when: "Коли потрібно мати кілька варіантів алгоритму і перемикатися між ними під час виконання.",
  py: `from abc import ABC, abstractmethod

class SortStrategy(ABC):
    @abstractmethod
    def sort(self, data): pass

class BubbleSort(SortStrategy):
    def sort(self, data):
        d = data[:]
        for i in range(len(d)):
            for j in range(len(d)-i-1):
                if d[j] > d[j+1]: d[j], d[j+1] = d[j+1], d[j]
        return d

class QuickSort(SortStrategy):
    def sort(self, data):
        if len(data) <= 1: return data
        p = data[len(data)//2]
        return (self.sort([x for x in data if x<p])
                + [x for x in data if x==p]
                + self.sort([x for x in data if x>p]))

class Sorter:
    def __init__(self, strategy: SortStrategy):
        self._strategy = strategy
    def set_strategy(self, s): self._strategy = s
    def sort(self, data): return self._strategy.sort(data)

s = Sorter(BubbleSort())
print(s.sort([3,1,4,1,5,9]))   # [1,1,3,4,5,9]
s.set_strategy(QuickSort())
print(s.sort([3,1,4,1,5,9]))   # [1,1,3,4,5,9]`,

  cpp: `#include <iostream>
#include <vector>
#include <algorithm>
using namespace std;

struct Strategy {
    virtual vector<int> sort(vector<int> v)=0;
    virtual ~Strategy()=default;
};

struct BubbleSort : Strategy {
    vector<int> sort(vector<int> v) override {
        for(int i=0;i<(int)v.size();i++)
          for(int j=0;j<(int)v.size()-i-1;j++)
            if(v[j]>v[j+1]) swap(v[j],v[j+1]);
        return v;
    }
};

struct StdSort : Strategy {
    vector<int> sort(vector<int> v) override {
        std::sort(v.begin(),v.end()); return v;
    }
};

class Sorter {
    Strategy* s;
public:
    Sorter(Strategy* strat): s(strat) {}
    void setStrategy(Strategy* strat) { s = strat; }
    vector<int> sort(vector<int> v) { return s->sort(v); }
};

int main() {
    BubbleSort bs; StdSort ss;
    Sorter sorter(&bs);
    for(int x: sorter.sort({3,1,4,1,5})) cout<<x<<" ";
    cout<<endl;
    sorter.setStrategy(&ss);
    for(int x: sorter.sort({3,1,4,1,5})) cout<<x<<" ";
}`,

  pas: `program StrategyDemo;

type
  TData = array of Integer;
  TStrategy = class
    function Sort(D: TData): TData; virtual; abstract;
  end;
  TBubbleSort = class(TStrategy)
    function Sort(D: TData): TData; override;
  end;
  TSorter = class
    FStrategy: TStrategy;
    constructor Create(S: TStrategy);
    function Sort(D: TData): TData;
  end;

function TBubbleSort.Sort(D: TData): TData;
var I, J, Tmp: Integer;
begin
  Result := Copy(D, 0, Length(D));
  for I:=0 to High(Result) do
    for J:=0 to High(Result)-I-1 do
      if Result[J] > Result[J+1] then
        begin Tmp:=Result[J]; Result[J]:=Result[J+1]; Result[J+1]:=Tmp; end;
end;

constructor TSorter.Create(S: TStrategy); begin FStrategy:=S; end;
function TSorter.Sort(D: TData): TData; begin Result:=FStrategy.Sort(D); end;

var S: TSorter; D, R: TData; I: Integer;
begin
  SetLength(D,5); D[0]:=3;D[1]:=1;D[2]:=4;D[3]:=1;D[4]:=5;
  S:=TSorter.Create(TBubbleSort.Create);
  R:=S.Sort(D);
  for I:=0 to High(R) do Write(R[I],' ');
  WriteLn;
  S.Free;
end.`
},

{
  name: "21. Template Method (Шаблонний метод)",
  intent: "Визначає скелет алгоритму в операції, відкладаючи деякі кроки на підкласи.",
  when: "Коли потрібно раз і назавжди зафіксувати незмінні частини алгоритму, залишивши підкласам реалізацію варіативних частин.",
  py: `from abc import ABC, abstractmethod

class DataProcessor(ABC):
    def process(self):        # Шаблонний метод
        data = self.read()
        parsed = self.parse(data)
        self.save(parsed)

    @abstractmethod
    def read(self): pass
    @abstractmethod
    def parse(self, data): pass

    def save(self, data):     # Хук — можна перевизначити
        print(f"Збережено: {data}")

class CSVProcessor(DataProcessor):
    def read(self):          return "1,2,3,4,5"
    def parse(self, data):   return list(map(int, data.split(',')))

class JSONProcessor(DataProcessor):
    def read(self):          return '{"values":[10,20,30]}'
    def parse(self, data):
        import json
        return json.loads(data)['values']

CSVProcessor().process()
JSONProcessor().process()`,

  cpp: `#include <iostream>
#include <vector>
#include <string>
using namespace std;

class DataProcessor {
public:
    void process() {        // template method
        auto data = read();
        auto parsed = parse(data);
        save(parsed);
    }
    virtual ~DataProcessor()=default;
protected:
    virtual string read()=0;
    virtual vector<int> parse(string data)=0;
    virtual void save(vector<int> d) {
        cout<<"Збережено: ";
        for(int x:d) cout<<x<<" ";
        cout<<endl;
    }
};

class CSVProcessor : public DataProcessor {
    string read() override { return "1,2,3,4,5"; }
    vector<int> parse(string s) override {
        vector<int> v; int n=0;
        for(char c:s+",") { if(c==','){ v.push_back(n);n=0;} else n=n*10+(c-'0'); }
        return v;
    }
};

int main() { CSVProcessor().process(); }`,

  pas: `program TemplateMethodDemo;

type
  TDataProcessor = class
    procedure Process;
    function  Read: string; virtual; abstract;
    function  Parse(const D: string): string; virtual; abstract;
    procedure Save(const D: string); virtual;
  end;
  TCSVProcessor = class(TDataProcessor)
    function Read: string; override;
    function Parse(const D: string): string; override;
  end;

procedure TDataProcessor.Process;
begin Save(Parse(Read)); end;
procedure TDataProcessor.Save(const D: string);
begin WriteLn('Збережено: ', D); end;

function TCSVProcessor.Read: string;  begin Result := '1,2,3,4,5'; end;
function TCSVProcessor.Parse(const D: string): string; begin Result := '['+D+']'; end;

begin
  TCSVProcessor.Create.Process;
end.`
},

{
  name: "22. Visitor (Відвідувач)",
  intent: "Описує операцію, що виконується над кожним елементом об'єктної структури. Дозволяє визначити нові операції, не змінюючи класи елементів.",
  when: "Коли потрібно виконати операцію над об'єктами різнорідних класів і не можна змінювати їх класи.",
  py: `from abc import ABC, abstractmethod

class Visitor(ABC):
    @abstractmethod
    def visit_circle(self, c): pass
    @abstractmethod
    def visit_square(self, s): pass

class Shape(ABC):
    @abstractmethod
    def accept(self, visitor): pass

class Circle(Shape):
    def __init__(self, r): self.r = r
    def accept(self, v): v.visit_circle(self)

class Square(Shape):
    def __init__(self, a): self.a = a
    def accept(self, v): v.visit_square(self)

class AreaVisitor(Visitor):
    def visit_circle(self, c):
        print(f"S коло = {3.14*c.r**2:.2f}")
    def visit_square(self, s):
        print(f"S квадрат = {s.a**2}")

shapes = [Circle(5), Square(4), Circle(3)]
av = AreaVisitor()
for s in shapes: s.accept(av)`,

  cpp: `#include <iostream>
#include <cmath>
using namespace std;

struct Circle; struct Square;
struct Visitor {
    virtual void visitCircle(Circle& c)=0;
    virtual void visitSquare(Square& s)=0;
    virtual ~Visitor()=default;
};

struct Shape { virtual void accept(Visitor& v)=0; virtual ~Shape()=default; };
struct Circle : Shape { double r; Circle(double r): r(r){}
    void accept(Visitor& v) override { v.visitCircle(*this); }
};
struct Square : Shape { double a; Square(double a): a(a){}
    void accept(Visitor& v) override { v.visitSquare(*this); }
};

struct AreaVisitor : Visitor {
    void visitCircle(Circle& c) override { printf("S коло = %.2f\n", M_PI*c.r*c.r); }
    void visitSquare(Square& s) override { printf("S квадрат = %.0f\n", s.a*s.a); }
};

int main() {
    Circle c(5); Square s(4);
    AreaVisitor av;
    c.accept(av); s.accept(av);
}`,

  pas: `program VisitorDemo;

type
  TCircle = class; TSquare = class;
  TVisitor = class
    procedure VisitCircle(C: TCircle); virtual; abstract;
    procedure VisitSquare(S: TSquare); virtual; abstract;
  end;
  TShape = class
    procedure Accept(V: TVisitor); virtual; abstract;
  end;
  TCircle = class(TShape)
    R: Double;
    constructor Create(Rad: Double);
    procedure Accept(V: TVisitor); override;
  end;
  TSquare = class(TShape)
    A: Double;
    constructor Create(Side: Double);
    procedure Accept(V: TVisitor); override;
  end;
  TAreaVisitor = class(TVisitor)
    procedure VisitCircle(C: TCircle); override;
    procedure VisitSquare(S: TSquare); override;
  end;

constructor TCircle.Create(Rad: Double); begin R:=Rad; end;
constructor TSquare.Create(Side: Double); begin A:=Side; end;

procedure TCircle.Accept(V: TVisitor); begin V.VisitCircle(Self); end;
procedure TSquare.Accept(V: TVisitor); begin V.VisitSquare(Self); end;

procedure TAreaVisitor.VisitCircle(C: TCircle);
begin WriteLn('S коло = ', 3.14*C.R*C.R:0:2); end;
procedure TAreaVisitor.VisitSquare(S: TSquare);
begin WriteLn('S квадрат = ', S.A*S.A:0:0); end;

var V: TAreaVisitor;
begin
  V:=TAreaVisitor.Create;
  TCircle.Create(5).Accept(V);
  TSquare.Create(4).Accept(V);
  V.Free;
end.`
},

{
  name: "23. Interpreter (Інтерпретатор)",
  intent: "Для заданої мови визначає представлення її граматики, а також інтерпретатор речень цієї мови.",
  when: "Коли потрібно інтерпретувати вирази деякої мови, яку можна представити як абстрактне синтаксичне дерево.",
  py: `from abc import ABC, abstractmethod

class Expression(ABC):
    @abstractmethod
    def interpret(self, ctx): pass

class Number(Expression):
    def __init__(self, val): self.val = val
    def interpret(self, ctx): return self.val

class Add(Expression):
    def __init__(self, l, r): self.l=l; self.r=r
    def interpret(self, ctx): return self.l.interpret(ctx)+self.r.interpret(ctx)

class Multiply(Expression):
    def __init__(self, l, r): self.l=l; self.r=r
    def interpret(self, ctx): return self.l.interpret(ctx)*self.r.interpret(ctx)

# (2 + 3) * 4 = 20
expr = Multiply(Add(Number(2), Number(3)), Number(4))
print(expr.interpret({}))   # 20`,

  cpp: `#include <iostream>
using namespace std;

struct Expr { virtual int eval()=0; virtual ~Expr()=default; };

struct Num : Expr { int v; Num(int v):v(v){}
    int eval() override { return v; }
};
struct Add : Expr { Expr *l,*r; Add(Expr* a,Expr* b):l(a),r(b){}
    int eval() override { return l->eval()+r->eval(); }
};
struct Mul : Expr { Expr *l,*r; Mul(Expr* a,Expr* b):l(a),r(b){}
    int eval() override { return l->eval()*r->eval(); }
};

int main() {
    // (2 + 3) * 4
    Expr* e = new Mul(new Add(new Num(2), new Num(3)), new Num(4));
    cout << e->eval() << endl;   // 20
}`,

  pas: `program InterpreterDemo;

type
  TExpr = class
    function Eval: Integer; virtual; abstract;
  end;
  TNum = class(TExpr)
    V: Integer;
    constructor Create(N: Integer);
    function Eval: Integer; override;
  end;
  TAdd = class(TExpr)
    L, R: TExpr;
    constructor Create(A, B: TExpr);
    function Eval: Integer; override;
  end;
  TMul = class(TExpr)
    L, R: TExpr;
    constructor Create(A, B: TExpr);
    function Eval: Integer; override;
  end;

constructor TNum.Create(N: Integer); begin V:=N; end;
function TNum.Eval: Integer; begin Result:=V; end;

constructor TAdd.Create(A, B: TExpr); begin L:=A; R:=B; end;
function TAdd.Eval: Integer; begin Result:=L.Eval+R.Eval; end;

constructor TMul.Create(A, B: TExpr); begin L:=A; R:=B; end;
function TMul.Eval: Integer; begin Result:=L.Eval*R.Eval; end;

var E: TMul;
begin
  { (2 + 3) * 4 = 20 }
  E := TMul.Create(TAdd.Create(TNum.Create(2),TNum.Create(3)),TNum.Create(4));
  WriteLn(E.Eval);
  E.Free;
end.`
}
];

// ─── BUILD DOCUMENT ───────────────────────────────────────────────────────────
function buildDoc() {
  const children = [];

  // Title page
  children.push(
    space(800),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "ПАТЕРНИ ПРОЄКТУВАННЯ", bold:true, size:56, color:C.h1bg, font:"Arial" })]
    }),
    space(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Design Patterns", bold:true, size:40, color:"888888", font:"Arial", italics:true })]
    }),
    space(400),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Навчальний посібник", size:28, color:"444444", font:"Arial" })]
    }),
    space(100),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Python  •  C++  •  Turbo Pascal", size:26, color:C.h2bg, font:"Arial", bold:true })]
    }),
    space(600),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Усі 23 патерни GoF (Gang of Four)", size:22, color:"666666", font:"Arial", italics:true })]
    }),
    pageBreak()
  );

  // Summary table
  children.push(
    h1("Зведена таблиця патернів"),
    space(120)
  );

  const tableData = [
    ["Група", "Патерн", "Ключова ідея"],
    ["Породжуючі","Singleton","Один екземпляр класу"],
    ["","Factory Method","Фабричний метод у підкласах"],
    ["","Abstract Factory","Родини пов'язаних об'єктів"],
    ["","Builder","Покрокове конструювання"],
    ["","Prototype","Клонування об'єктів"],
    ["Структурні","Adapter","Перетворення інтерфейсу"],
    ["","Bridge","Відокремлення абстракції і реалізації"],
    ["","Composite","Дерево частина-ціле"],
    ["","Decorator","Динамічне додавання обов'язків"],
    ["","Facade","Спрощений інтерфейс до підсистеми"],
    ["","Flyweight","Спільний стан для дрібних об'єктів"],
    ["","Proxy","Замісник з контрольованим доступом"],
    ["Поведінкові","Chain of Responsibility","Ланцюжок обробників"],
    ["","Command","Запит як об'єкт + undo"],
    ["","Iterator","Послідовний перебір колекції"],
    ["","Mediator","Централізована взаємодія"],
    ["","Memento","Збереження та відновлення стану"],
    ["","Observer","Підписка на зміни стану"],
    ["","State","Поведінка залежить від стану"],
    ["","Strategy","Взаємозамінні алгоритми"],
    ["","Template Method","Скелет алгоритму у базовому класі"],
    ["","Visitor","Нові операції без зміни класів"],
    ["","Interpreter","Граматика і інтерпретація мови"],
  ];

  const tBorder = border1("AAAAAA");
  const tBorders = {top:tBorder,bottom:tBorder,left:tBorder,right:tBorder};
  const colW = [2200,2600,4560];

  const tableRows = tableData.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: tBorders,
      width: { size: colW[ci], type: WidthType.DXA },
      shading: { fill: ri===0 ? C.tablHead : (ri%2===0 ? "F5F9FF" : C.white), type: ShadingType.CLEAR },
      margins: { top:80, bottom:80, left:120, right:120 },
      children: [new Paragraph({
        children: [new TextRun({ text: cell, bold: ri===0, size: ri===0?20:20, font:"Arial" })]
      })]
    }))
  }));

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colW,
    rows: tableRows
  }));
  children.push(pageBreak());

  // Patterns
  let currentGroup = "";
  for (const p of patterns) {
    if (p.group && p.group !== currentGroup) {
      currentGroup = p.group;
      children.push(h1(currentGroup), space(160));
    }

    children.push(
      h2(p.name),
      space(80),
      h3("Призначення"),
      para(p.intent),
      h3("Коли застосовувати"),
      para(p.when),
      space(80)
    );
    children.push(pyCode(p.py));
    children.push(cppCode(p.cpp));
    children.push(pasCode(p.pas));
    children.push(pageBreak());
  }

  // Remove last page break
  children.pop();

  return new Document({
    styles: {
      default: { document: { run: { font:"Arial", size:22 } } },
      paragraphStyles: [
        { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:40, bold:true, font:"Arial", color:C.h1fg },
          paragraph:{ spacing:{before:360,after:240}, outlineLevel:0 } },
        { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:32, bold:true, font:"Arial", color:C.h2fg },
          paragraph:{ spacing:{before:300,after:200}, outlineLevel:1 } },
        { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:26, bold:true, font:"Arial", color:C.h3fg },
          paragraph:{ spacing:{before:240,after:160}, outlineLevel:2 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width:11906, height:16838 },
          margin: { top:1080, right:1080, bottom:1080, left:1080 }
        }
      },
      children
    }]
  });
}

Packer.toBuffer(buildDoc()).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/design_patterns_ua.docx", buf);
  console.log("Done!");
}).catch(console.error);
