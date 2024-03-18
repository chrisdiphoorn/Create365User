
Using Namespace System
Using Namespace System.Drawing
Using Namespace System.Windows
Using Namespace System.Windows.Forms
Using Namespace System.Management.Automation

# MTA - Multi-Threaded Apartment
# STA - Single Threaded Apartment - for TextBox, ComboBox, DataGrid, Form Top = True
# PowerShell.exe –STA
$PowershellMode = $host.Runspace.ApartmentState

# Update powershell to 7
# iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"

#if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
##Requires -RunAsAdministrator

# You will need to be able to add new powershell modules if they are not already installed.
$ExecPolicy = (Get-ExecutionPolicy -Scope CurrentUser)
if($ExecPolicy -ne 'RemoteSigned') {
	try {
		Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
	} catch {
		#write-Host "This script needs to be run with the Powershell Excution policy RemoteSigned"
		#write-host ""
		#pause
		#exit
	}
}

function Get-CurrentPath
	{
		$currentPath = $PSScriptRoot                                                                                                     # AzureDevOps, Powershell
		if (!$currentPath) { $currentPath = Split-Path $pseditor.GetEditorContext().CurrentFile.Path -ErrorAction SilentlyContinue }     # VSCode
		if (!$currentPath) { $currentPath = Split-Path $psISE.CurrentFile.FullPath -ErrorAction SilentlyContinue }                       # PsISE
		return $currentPath + '\'
	}

$Default = $null
$cfolder = Get-CurrentPath 
$IniFile = "$($cFolder)Create365User.ini"
# Read the INI File if it exists
IF([System.IO.File]::Exists($IniFile) -eq $true) {
	$Default = Get-Content $IniFile | ConvertFrom-StringData
}

$script:Tenant = ''
$script:AzureClientApp = ''
$script:AzureClientPassword = ''
$script:ConnectSPOServiceUser = ''
$script:ConnectSPOServicePassword = ''
$script:ActiveDirectoryUsername = ''
$script:ActiveDirectoryPassword = ''

$OUPath = ""

#Default Address for an Australian User
$AUpostalcode = ""
$AUstate = ""
$AUstreetaddress = ""
$AUcity = ""
$AUcountry = ""

#Default Address for a Philippines User
$Secondpostalcode = ""
$SecondState = ""
$SecondStreetAddress = ""
$SecondCity = ""
$SecondCountry = ""

#Sharepoint Details
$SharepointGroups = @()
$SPAdminSite = ''
$SPSite = ''

# License SKUid Details for Tenant: 
$E1 = ""
$E3 = ""
$E5 = "" #None Purchased.
$BP = ""
$BE = ""
$Defender = ""
$IntuneSuite = ""
$VisioClient = ""
$NonProfitPortal = ""

$company = ''
$WWW = ''
$EmailDomain = ''
$SMTPServer = ''
$SendEmailAddress = ''

$AllStaffSigGroup = ''
$OutlookSigGroup = ''

if($Default -ne $null) {
	if($Default.Tenant) { $script:Tenant = $Default.Tenant }
	if($Default.AzureClientApp) { $script:AzureClientApp = $Default.AzureClientApp }
	if($Default.AzureClientPassword) { $script:AzureClientPassword = $Default.AzureClientPassword }
	if($Default.ConnectSPOServiceUser) { $script:ConnectSPOServiceUser = $Default.ConnectSPOServiceUser }
	if($Default.ConnectSPOServicePassword ) { $script:ConnectSPOServicePassword = $Default.ConnectSPOServicePassword }
	if($Default.ActiveDirectoryUsername) { $script:ActiveDirectoryUsername = $Default.ActiveDirectoryUsername }
	if($Default.ActiveDirectoryPassword) { $script:ActiveDirectoryPassword = $Default.ActiveDirectoryPassword }
	if($Default.OUPath) { $OUPath = $Default.OUPath }
	if($Default.SharepointGroups) {
		$split = $Default.SharepointGroups -Split ","
		$SharepointGroups = @()
		foreach($name in $Split) {
			$name = $name.replace("'","").trim()
			$SharepointGroups += $name
		}
	}
	if($Default.SPAdminSite) { $SPAdminSite = $Default.SPAdminSite} 
	if($Default.SPSite) { $SPSite = $Default.SPSite} 
	
	if ($Default.LicenseE1) { $E1 = $Default.LicenseE1}
	if ($Default.LicenseE3) { $E3 = $Default.LicenseE3}
	if ($Default.LicenseE5) { $E5 = $Default.LicenseE5}
	if ($Default.LicenseBP) { $BP = $Default.LicenseBP}
	if ($Default.LicenseBE) { $BE = $Default.LicenseBE}
	if ($Default.LicenseDefender) { $Defender = $Default.LicenseDefender}
	if ($Default.LicenseInTuneSuite) { $InTuneSuite = $Default.LicenseInTuneSuite}
	if ($Default.LicenseVisioClient) { $VisioClient = $Default.LicenseVisioClient}
	if ($Default.LicenseNonProfitPortal) { $NonProfitPortal = $Default.LicenseNonProfitPortal}
	
	if($Default.DefaultAUPostalCode) { $AUPostalCode= $Default.DefaultAUPostalCode}
	if($Default.DefaultAUstate) { $AUstate= $Default.DefaultAUstate}
	if($Default.DefaultAUstreetaddress) { $AUstreetaddress = $Default.DefaultAUstreetaddress}
	if($Default.DefaultAUcity) { $AUcity = $Default.DefaultAUcity}
	if($Default.DefaultAUcountry) { $AUcountry = $Default.DefaultAUcountry}
	
	if($Default.DefaultSecondPostalCode) { $Secondpostalcode= $Default.DefaultSecondPostalCode}
	if($Default.DefaultSecondstate) { $SecondState= $Default.DefaultSecondstate}
	if($Default.DefaultSecondstreetaddress) { $SecondStreetAddress = $Default.DefaultSecondstreetaddress}
	if($Default.DefaultSecondcity) { $SecondCity = $Default.DefaultSecondcity}
	if($Default.DefaultSecondcountry) { $SecondCountry = $Default.DefaultSecondcountry}
	if($Default.Company) { $company = $Default.Company}
	if($Default.WWW) { $WWW = $Default.WWW}
	if($Default.Domain) { $EmailDomain = $Default.Domain}
	if($Default.SMTPServer) { $SMTPServer = $Default.SMTPServer}
	if($Default.SendEmail) { $SendEmailAddress = $Default.SendEmail}
	if($Default.GroupOutlook) { $OutlookSigGroup = $Default.GroupOutlook}
	if($Default.GroupAllStaff) { $AllStaffSigGroup = $Default.GroupAllStaff}
	
}


# KeyPress codes that wont work in a textbox
$invalidKeys =@('next','multiply','OemBackslash','OemClear','OemCloseBrackets','Oemcomma', 'OemOpenBrackets','OemPipe','Oemplus','OemQuestion','OemQuotes','OemSemicolon','Oemtilde','Oem5','Oem6','d3','d4','d5','d6','d7','d8','d9','divide')

# Password Dictionaly of words
$nouns = @('able','account','achieve','achiever','acoustics','act','action','activity','actor','addition','adjustment','advertisement','advice','aftermath','afternoon','afterthought','agreement','air','airplane','airport','alarm','alley','amount','amusement','anger','angle','animal','answer','ant','ants','apparatus','apparel','apple','apples','appliance','approval','arch','argument','arithmetic','arm','army','art','attack','attempt','attention','attraction','aunt','authority','babies','baby','back','badge','bag','bait','balance','ball','balloon','balls','banana','band','base','baseball','basin','basket','basketball','bat','bath','battle','bead','beam','bean','bear','bears','beast','bed','bedroom','beds','bee','beef','beetle','beggar','beginner','behavior','belief','believe','bell','bells','berry','bike','bikes','bird','birds','birth','birthday','bit','bite','blade','blood','blow','board','boat','boats','body','bomb','bone','book','books','boot','border','bottle','boundary','box','boy','boys','brain','brake','branch','brass','bread','breakfast','breath','brick','bridge','brother','brothers','brush','bubble','bucket','building','bulb','bun','burn','burst','bushes','business','butter','button','cabbage','cable','cactus','cake','cakes','calculator','calendar','camera','camp','can','cannon','canvas','cap','caption','car','card','care','carpenter','carriage','cars','cart','cast','cat','cats','cattle','cause','cave','celery','cellar','cemetery','cent','chain','chair','chairs','chalk','chance','change','channel','cheese','cherries','cherry','chess','chicken','chickens','children','chin','church','circle','clam','class','clock','clocks','cloth','cloud','clouds','clover','club','coach','coal','coast','coat','cobweb','coil','collar','color','comb','comfort','committee','company','comparison','competition','condition','connection','control','cook','copper','copy','cord','cork','corn','cough','country','cover','cow','cows','crack','cracker','crate','crayon','cream','creator','creature','credit','crib','crime','crook','crow','crowd','crown','crush','cry','cub','cup','current','curtain','curve','cushion','dad','daughter','day','death','debt','decision','deer','degree','design','desire','desk','destruction','detail','development','digestion','dime','dinner','dinosaurs','direction','dirt','discovery','discussion','disease','disgust','distance','distribution','division','dock','doctor','dog','dogs','doll','dolls','donkey','door','downtown','drain','drawer','dress','drink','driving','drop','drug','drum','duck','ducks','dust','ear','earth','earthquake','edge','education','effect','egg','eggnog','eggs','elbow','end','engine','error','event','example','exchange','existence','expansion','experience','expert','eye','eyes','face','fact','fairies','fall','family','fan','fang','farm','farmer','father','faucet','fear','feast','feather','feeling','feet','fiction','field','fifth','fight','finger','fire','fireman','fish','flag','flame','flavor','flesh','flight','flock','floor','flower','flowers','fly','fog','fold','food','foot','force','fork','form','fowl','frame','friction','friend','friends','frog','frogs','front','fruit','fuel','furniture','galley','game','garden','gate','geese','ghost','giants','giraffe','girl','girls','glass','glove','glue','goat','gold','goldfish','good-bye','goose','government','governor','grade','grain','grandfather','grandmother','grape','grass','grip','ground','group','growth','guide','guitar','gun','hair','haircut','hall','hammer','hand','hands','harbor','harmony','hat','hate','head','health','hearing','heart','heat','help','hen','hill','history','hobbies','hole','holiday','home','honey','hook','hope','horn','horse','horses','hose','hospital','hot','hour','house','houses','humor','hydrant','ice','icicle','idea','impulse','income','increase','industry','ink','insect','instrument','insurance','interest','invention','iron','island','jail','jam','jar','jeans','jelly','jellyfish','jewel','join','joke','journey','judge','juice','jump','kettle','key','kick','kiss','kite','kitten','kittens','kitty','knee','knife','knot','knowledge','laborer','lace','ladybug','lake','lamp','land','language','laugh','lawyer','lead','leaf','learning','leather','leg','legs','letter','letters','lettuce','level','library','lift','light','limit','line','linen','lip','liquid','list','lizards','loaf','lock','locket','look','loss','love','low','lumber','lunch','lunchroom','machine','magic','maid','mailbox','man','manager','map','marble','mark','market','mask','mass','match','meal','measure','meat','meeting','memory','men','metal','mice','middle','milk','mind','mine','minister','mint','minute','mist','mitten','mom','money','monkey','month','moon','morning','mother','motion','mountain','mouth','move','muscle','music','nail','name','nation','neck','need','needle','nerve','nest','net','news','night','noise','north','nose','note','notebook','number','nut','oatmeal','observation','ocean','offer','office','oil','operation','opinion','orange','oranges','order','organization','ornament','oven','owl','owner','page','pail','pain','paint','pan','pancake','paper','parcel','parent','park','part','partner','party','passenger','paste','patch','payment','peace','pear','pen','pencil','person','pest','pet','pets','pickle','picture','pie','pies','pig','pigs','pin','pipe','pizzas','place','plane','planes','plant','plantation','plants','plastic','plate','play','playground','pleasure','plot','plough','pocket','point','poison','police','polish','pollution','popcorn','porter','position','pot','potato','powder','power','price','print','prison','process','produce','profit','property','prose','protest','pull','pump','punishment','purpose','push','quarter','quartz','queen','question','quicksand','quiet','quill','quilt','quince','quiver','rabbit','rabbits','rail','railway','rain','rainstorm','rake','range','rat','rate','ray','reaction','reading','reason','receipt','recess','record','regret','relation','religion','representative','request','respect','rest','reward','rhythm','rice','riddle','rifle','ring','rings','river','road','robin','rock','rod','roll','roof','room','root','rose','route','rub','rule','run','sack','sail','salt','sand','scale','scarecrow','scarf','scene','scent','school','science','scissors','screw','sea','seashore','seat','secretary','seed','selection','self','sense','servant','shade','shake','shame','shape','sheep','sheet','shelf','ship','shirt','shock','shoe','shoes','shop','show','side','sidewalk','sign','silk','silver','sink','sister','sisters','size','skate','skin','skirt','sky','slave','sleep','sleet','slip','slope','smash','smell','smile','smoke','snail','snails','snake','snakes','sneeze','snow','soap','society','sock','soda','sofa','son','song','songs','sort','sound','soup','space','spade','spark','spiders','sponge','spoon','spot','spring','spy','square','squirrel','stage','stamp','star','start','statement','station','steam','steel','stem','step','stew','stick','sticks','stitch','stocking','stomach','stone','stop','store','story','stove','stranger','straw','stream','street','stretch','string','structure','substance','sugar','suggestion','suit','summer','sun','support','surprise','sweater','swim','swing','system','table','tail','talk','tank','taste','tax','teaching','team','teeth','temper','tendency','tent','territory','test','texture','theory','thing','things','thought','thread','thrill','throat','throne','thumb','thunder','ticket','tiger','time','tin','title','toad','toe','toes','tomatoes','tongue','tooth','toothbrush','toothpaste','top','touch','town','toy','toys','trade','trail','train','trains','tramp','transport','tray','treatment','tree','trees','trick','trip','trouble','trousers','truck','trucks','tub','turkey','turn','twig','twist','umbrella','uncle','underwear','unit','use','vacation','value','van','vase','vegetable','veil','vein','verse','vessel','vest','view','visitor','voice','volcano','volleyball','voyage','walk','wall','war','wash','waste','watch','water','wave','waves','wax','way','wealth','weather','week','weight','wheel','whip','whistle','wilderness','wind','window','wine','wing','winter','wire','wish','woman','women','wood','wool','word','work','worm','wound','wren','wrench','wrist','writer','writing','yak','yam','yard','yarn','year','yoke','zebra','zephyr','zinc','zipper','zoo')
$verbs = @('abide','accelerate','accept','accomplish','achieve','acquire','acted','activate','adapt','add','address','administer','admire','admit','adopt','advise','afford','agree','alert','alight','allow','altered','amuse','analyze','announce','annoy','answer','anticipate','apologize','appear','applaud','applied','appoint','appraise','appreciate','approve','arbitrate','argue','arise','arrange','arrest','arrive','ascertain','ask','assemble','assess','assist','assure','attach','attack','attain','attempt','attend','attract','audited','avoid','awake','back','bake','balance','ban','bang','bare','bat','bathe','battle','be','beam','bear','beat','become','beg','begin','behave','behold','belong','bend','beset','bet','bid','bind','bite','bleach','bleed','bless','blind','blink','blot','blow','blush','boast','boil','bolt','bomb','book','bore','borrow','bounce','bow','box','brake','branch','break','breathe','breed','brief','bring','broadcast','bruise','brush','bubble','budget','build','bump','burn','burst','bury','bust','buy','buze','calculate','call','camp','care','carry','carve','cast','catalog','catch','cause','challenge','change','charge','chart','chase','cheat','check','cheer','chew','choke','choose','chop','claim','clap','clarify','classify','clean','clear','cling','clip','close','clothe','coach','coil','collect','color','comb','come','command','communicate','compare','compete','compile','complain','complete','compose','compute','conceive','concentrate','conceptualize','concern','conclude','conduct','confess','confront','confuse','connect','conserve','consider','consist','consolidate','construct','consult','contain','continue','contract','control','convert','coordinate','copy','correct','correlate','cost','cough','counsel','count','cover','crack','crash','crawl','create','creep','critique','cross','crush','cry','cure','curl','curve','cut','cycle','dam','damage','dance','dare','deal','decay','deceive','decide','decorate','define','delay','delegate','delight','deliver','demonstrate','depend','describe','desert','deserve','design','destroy','detail','detect','determine','develop','devise','diagnose','dig','direct','disagree','disappear','disapprove','disarm','discover','dislike','dispense','display','disprove','dissect','distribute','dive','divert','divide','do','double','doubt','draft','drag','drain','dramatize','draw','dream','dress','drink','drip','drive','drop','drown','drum','dry','dust','dwell','earn','eat','edited','educate','eliminate','embarrass','employ','empty','enacted','encourage','end','endure','enforce','engineer','enhance','enjoy','enlist','ensure','enter','entertain','escape','establish','estimate','evaluate','examine','exceed','excite','excuse','execute','exercise','exhibit','exist','expand','expect','expedite','experiment','explain','explode','express','extend','extract','face','facilitate','fade','fail','fancy','fasten','fax','fear','feed','feel','fence','fetch','fight','file','fill','film','finalize','finance','find','fire','fit','fix','flap','flash','flee','fling','float','flood','flow','flower','fly','fold','follow','fool','forbid','force','forecast','forego','foresee','foretell','forget','forgive','form','formulate','forsake','frame','freeze','frighten','fry','gather','gaze','generate','get','give','glow','glue','go','govern','grab','graduate','grate','grease','greet','grin','grind','grip','groan','grow','guarantee','guard','guess','guide','hammer','hand','handle','handwrite','hang','happen','harass','harm','hate','haunt','head','heal','heap','hear','heat','help','hide','hit','hold','hook','hop','hope','hover','hug','hum','hunt','hurry','hurt','hypothesize','identify','ignore','illustrate','imagine','implement','impress','improve','improvise','include','increase','induce','influence','inform','initiate','inject','injure','inlay','innovate','input','inspect','inspire','install','institute','instruct','insure','integrate','intend','intensify','interest','interfere','interlay','interpret','interrupt','interview','introduce','invent','inventory','investigate','invite','irritate','itch','jail','jam','jog','join','joke','judge','juggle','jump','justify','keep','kept','kick','kill','kiss','kneel','knit','knock','knot','know','label','land','last','laugh','launch','lay','lead','lean','leap','learn','leave','lecture','led','lend','let','level','license','lick','lie','lifted','light','lighten','like','list','listen','live','load','locate','lock','log','long','look','lose','love','maintain','make','man','manage','manipulate','manufacture','map','march','mark','market','marry','match','mate','matter','mean','measure','meddle','mediate','meet','melt','melt','memorize','mend','mentor','milk','mine','mislead','miss','misspell','mistake','misunderstand','mix','moan','model','modify','monitor','moor','motivate','mourn','move','mow','muddle','mug','multiply','murder','nail','name','navigate','need','negotiate','nest','nod','nominate','normalize','note','notice','number','obey','object','observe','obtain','occur','offend','offer','officiate','open','operate','order','organize','oriented','originate','overcome','overdo','overdraw','overflow','overhear','overtake','overthrow','owe','own','pack','paddle','paint','park','part','participate','pass','paste','pat','pause','pay','peck','pedal','peel','peep','perceive','perfect','perform','permit','persuade','phone','photograph','pick','pilot','pinch','pine','pinpoint','pioneer','place','plan','plant','play','plead','please','plug','point','poke','polish','pop','possess','post','pour','practice','praised','pray','preach','precede','predict','prefer','prepare','prescribe','present','preserve','preset','preside','press','pretend','prevent','prick','print','process','procure','produce','profess','program','progress','project','promise','promote','proofread','propose','protect','prove','provide','publicize','pull','pump','punch','puncture','punish','purchase','push','put','qualify','question','queue','quit','race','radiate','rain','raise','rank','rate','reach','read','realign','realize','reason','receive','recognize','recommend','reconcile','record','recruit','reduce','refer','reflect','refuse','regret','regulate','rehabilitate','reign','reinforce','reject','rejoice','relate','relax','release','rely','remain','remember','remind','remove','render','reorganize','repair','repeat','replace','reply','report','represent','reproduce','request','rescue','research','resolve','respond','restored','restructure','retire','retrieve','return','review','revise','rhyme','rid','ride','ring','rinse','rise','risk','rob','rock','roll','rot','rub','ruin','rule','run','rush','sack','sail','satisfy','save','saw','say','scare','scatter','schedule','scold','scorch','scrape','scratch','scream','screw','scribble','scrub','seal','search','secure','see','seek','select','sell','send','sense','separate','serve','service','set','settle','sew','shade','shake','shape','share','shave','shear','shed','shelter','shine','shiver','shock','shoe','shoot','shop','show','shrink','shrug','shut','sigh','sign','signal','simplify','sin','sing','sink','sip','sit','sketch','ski','skip','slap','slay','sleep','slide','sling','slink','slip','slit','slow','smash','smell','smile','smite','smoke','snatch','sneak','sneeze','sniff','snore','snow','soak','solve','soothe','soothsay','sort','sound','sow','spare','spark','sparkle','speak','specify','speed','spell','spend','spill','spin','spit','split','spoil','spot','spray','spread','spring','sprout','squash','squeak','squeal','squeeze','stain','stamp','stand','stare','start','stay','steal','steer','step','stick','stimulate','sting','stink','stir','stitch','stop','store','strap','streamline','strengthen','stretch','stride','strike','string','strip','strive','stroke','structure','study','stuff','sublet','subtract','succeed','suck','suffer','suggest','suit','summarize','supervise','supply','support','suppose','surprise','surround','suspect','suspend','swear','sweat','sweep','swell','swim','swing','switch','symbolize','synthesize','systemize','tabulate','take','talk','tame','tap','target','taste','teach','tear','tease','telephone','tell','tempt','terrify','test','thank','thaw','think','thrive','throw','thrust','tick','tickle','tie','time','tip','tire','touch','tour','tow','trace','trade','train','transcribe','transfer','transform','translate','transport','trap','travel','tread','treat','tremble','trick','trip','trot','trouble','troubleshoot','trust','try','tug','tumble','turn','tutor','twist','type','undergo','understand','undertake','undress','unfasten','unify','unite','unlock','unpack','untidy','update','upgrade','uphold','upset','use','utilize','vanish','verbalize','verify','vex','visit','wail','wait','wake','walk','wander','want','warm','warn','wash','waste','watch','water','wave','wear','weave','wed','weep','weigh','welcome','wend','wet','whine','whip','whirl','whisper','whistle','win','wind','wink','wipe','wish','withdraw','withhold','withstand','wobble','wonder','work','worry','wrap','wreck','wrestle','wriggle','wring','write','xray','yawn','yell','zip','zoom')


# Create the HTML CSS for the Email Table
$Head = @'
<div>
<style>
   
  body {
    font-family: Arial;
    font-size: 8pt;
    color: #4C607B;
    }
  
   table { 
    cellpadding: 0;
    callspacing:0;
    margin: 0px;
    padding: 5px;
    border-spacing: 0px;
    border: 1px solid #D3D3D3;  
    border-collapse: collapse;
    font-size: 1.2em;
    text-align: left;
    }
  
  th {     
      border: 1px solid #D3D3D3; 
      padding: 5px;  
      background-color: #003366;
      color: #ffffff;
     }
  
  td {
      border: 1px solid #D3D3D3; 
      padding: 5px;
      color: #000000;
    }
 tr:nth-child(even) {background-color: #f2f2f2;}
     
  
</style>
</div>

'@

# C# Code to produce a Shadow under the form - Does not work?
$Shadow = @'

using System;
using System.Windows;
using System.Windows.Forms;

namespace Program
{
    public partial class Shadow: Form

    {
        protected override CreateParams CreateParams
        {
            get
            {
                const int CS_DROPSHADOW = 0x20000;
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }    
    }
}
'@

try {
    Add-Type -TypeDefinition $Shadow -Language CSharp -WarningAction SilentlyContinue -ReferencedAssemblies System, System.Windows, System.Windows.Forms, System.ComponentModel.Primitives
}
catch {}

# C# Code for creating round buttons using Win32Helpers
$code = @'
[System.Runtime.InteropServices.DllImport("gdi32.dll")]
public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
'@
try {
    $Win32Helpers = Add-Type -MemberDefinition $code -Name "Win32Helpers" -WarningAction SilentlyContinue -PassThru
} catch { }

# Functions to get Powershell Script Information
function PSCommandPath() { return $PSCommandPath }
function ScriptName() { return $MyInvocation.ScriptName }
function MyCommandName() { return $MyInvocation.MyCommand.Name }
function MyCommandDefinition() {
    # Begin of MyCommandDefinition()
    # Note: ouput of this script shows the contents of this function, not the execution result
    return $MyInvocation.MyCommand.Definition
    # End of MyCommandDefinition()
}
function MyInvocationPSCommandPath() { return $MyInvocation.PSCommandPath }

# The name of the powershell script that is curently running
$ScriptName = "$($MyInvocation.MyCommand.Name)"

# Add the .NET Functions used to create WinForms
try {
        Add-Type -AssemblyName System.Windows.Forms -WarningAction SilentlyContinue
        Add-Type -AssemblyName System.Drawing -WarningAction SilentlyContinue
        Add-Type -AssemblyName PresentationCore, PresentationFramework -WarningAction SilentlyContinue

    } catch {
        Write-Error -TargetObject $_ -Message "Exception encountered during Environment Setup." -Category ResourceUnavailable
        exit
    }

# Enable Windows Forms Visual Styles - Picks up the themed from the operating system
[System.Windows.Forms.Application]::EnableVisualStyles() 


# Set up the GUI Variables ***************************************************************************************

# Font styles are: Regular, Bold, Italic, Underline, Strikeout
$Font12 = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$Font10B = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$Font10 = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$Font9 = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Regular)
$Font8 = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Regular)
$Font14B = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$Tick_Font = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)

# Create Dark Background Colours 
$BackColor = [System.Drawing.Color]::FromArgb(42, 43, 47)
$ForeColor = [System.Drawing.Color]::FromKnownColor("ControlLight")

$ButtonFont = $Font10B
$ButtonBorderColor = [System.Drawing.Color]::FromArgb(163, 42, 64)
$ButtonColor = [System.Drawing.Color]::FromKnownColor("Black")
$ButtonMouseDownColor = [System.Drawing.Color]::FromArgb(237,56,89)
$ButtonMouseOverColor = [System.Drawing.Color]::FromArgb(42, 43, 47)

$ButtonBorderColor = [System.Drawing.Color]::FromArgb(163, 42, 64)
$ButtonColor = [System.Drawing.Color]::FromArgb(163, 42, 64)
$ButtonMouseDownColor = [System.Drawing.Color]::FromArgb(237,56,89)
$ButtonMouseOverColor = [System.Drawing.Color]::FromArgb(197, 35,65)

$TextBoxBackColor = [System.Drawing.Color]::Black
$TextBoxFont = $Font9
$Form_Font = $Font10

$global:ADusers = New-Object -TypeName PsObject
$global:UPNS = New-Object -TypeName PsObject

$global:SharedMailboxes = New-Object -TypeName PsObject

# Create an array to hold all gui objects
$Script:GUIFormObjectList = @() 

# Create an array to hold all gui starting enabled objects - #A Reset will put them back to these values.
$Script:GUIEnabledObjectList = @() 

# Create new Loading Form - using the shadow function.
$formLoading = New-Object Program.Shadow
$labelLoading = New-Object 'System.Windows.Forms.Label'
$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
  
   
$Form_StateCorrection_Load =
  {
    #Correct the initial state of the form to prevent the .Net maximized form issue;
    $formLoading.WindowState = $InitialFormWindowState;
    $hrgn = $Win32Helpers::CreateRoundRectRgn(0,0,$formLoading.Width, $formLoading.Height, 4,4);
    $formLoading.Region = [Region]::FromHrgn($hrgn);
}
  
$Form_Cleanup_FormClosed =
{
    #Remove all event handlers from the controls
    try
    {
      [void] $formLoading.remove_Load($Form_StateCorrection_Load);
      [void] $formLoading.remove_FormClosed($Form_Cleanup_FormClosed);
    } catch [Exception] { }
}
  
$formloading.Font = $Font14B

$ScriptText = " Loading $($ScriptName) - Please wait... "
$aSize = [System.Windows.Forms.TextRenderer]::MeasureText($ScriptText, $Font14B)
[int]$getWidth = $aSize.width+ 4 
  
[void]$formLoading.Controls.Add($labelLoading)

$formLoading.BackColor = $BackColor
$formLoading.ForeColor = $ForeColor

$formLoading.ControlBox = $False
$formLoading.Cursor = 'AppStarting'
#$formLoading.FormBorderStyle = 'FixedToolWindow'
$formLoading.FormBorderStyle = 'None'
$formLoading.Name = "Splash"
$formLoading.ShowIcon = $False
$formLoading.ShowInTaskbar = $False
$formLoading.StartPosition = 'CenterScreen'
$formLoading.Text = ""
$formLoading.AutoSize = $true
$formLoading.ClientSize = "$($getWidth), 41"
 
# Create a label to show on the form
$labelLoading.Location = '5, 5'
$labelLoading.Size = "$getWidth, 36"
$labelLoading.TabIndex = 0
$labelLoading.Text = $ScriptText

[void]$formLoading.ResumeLayout()
[System.Windows.Forms.Application]::DoEvents()

$formLoading_Shown = {
    $this.Activate();
}

# Activate the form after it has been painted
[void]$formLoading.Add_Shown( $formLoading_Shown )
    
#Save the initial state of the form
$InitialFormWindowState = $formLoading.WindowState

#Init the OnLoad event to correct the initial state of the form
[void]$formLoading.add_Load($Form_StateCorrection_Load)

#Clean up the control events
[void]$formLoading.add_FormClosed($Form_Cleanup_FormClosed)

#Show the Form
[void]$formLoading.Show((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
[System.Windows.Forms.Application]::DoEvents()

# Set a Tooltip Text when hovering on a Form Object
Function Set-Tooltip
{
    param(
        $Control,
        [string]$StrTooltip,
        [switch]$Show
    )

    $ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    $ToolTip.IsBalloon = $true
    $ToolTip.InitialDelay = 500
    $ToolTip.ReshowDelay = 500
    $ToolTip.SetToolTip($Control, $StrTooltip)

    if($Show)
    {
        $ToolTip.Show($StrTooltip, $Control, 1500)
    }
}

	
# See https://www.powershellgallery.com/ for module and version info
Function Install-ModuleIfNotInstalled(

		[string] [Parameter(Mandatory = $true)] $moduleName,
		[string] $minimalVersion, 
		[boolean] [Parameter(mandatory=$false)] $AllowPrerelease = $False
		)
	{
    $module = Get-Module -Name $moduleName -ListAvailable | Where-Object { $null -eq $minimalVersion -or $minimalVersion -lt $_.Version } | Select-Object -Last 1

    if ($null -ne $module) {
         Write-Verbose ('Module {0} (v{1}) is available.' -f $moduleName, $module.Version)
		 if($module.Version -eq $minimalVersion) {
			 Update-Module -Name $moduleName -RequiredVersion $minimalVersion -Force 
		 }
    } else {
        Import-Module -Name 'PowershellGet'
        $installedModule = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue
        if ($null -ne $installedModule) {
            Write-Verbose ('Module [{0}] (v {1}) is installed.' -f $moduleName, $installedModule.Version)
        }
        if ($null -eq $installedModule -or ($null -ne $minimalVersion -and $installedModule.Version -lt $minimalVersion)) {
            Write-Verbose ('Module {0} min.vers {1}: not installed; check if nuget v2.8.5.201 or later is installed.' -f $moduleName, $minimalVersion)
            #First check if package provider NuGet is installed. Incase an older version is installed the required version is installed explicitly
            if ((Get-PackageProvider -Name NuGet -Force).Version -lt '2.8.5.201') {
                Write-Warning ('Module {0} min.vers {1}: Install nuget!' -f $moduleName, $minimalVersion)
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
            }        
            $optionalArgs = New-Object -TypeName Hashtable
            if ($null -ne $minimalVersion) {
                $optionalArgs['RequiredVersion'] = $minimalVersion
            }  
            
            Write-Warning ('Install module {0} (version [{1}]) within scope of the current user.' -f $moduleName, $minimalVersion)
            if ( $AllowPrerelease -eq $False) {
                Install-Module -Name $moduleName @optionalArgs -Scope CurrentUser -Force -Verbose
            } else {
               Install-Module -Name $moduleName @optionalArgs -Scope CurrentUser -Force -Verbose -AllowPrerelease
            }
        } # Nuget Install
    } # Module Install
}

#Get a strong password using the Dinopass API
Function Get-Password-DinoPass()
{
    $dinopass = ""
    try
    {
        $dinopass = (Invoke-WebRequest -Uri 'https://www.dinopass.com/password/strong' -Method GET)
    } catch { }
    Return $dinopass
}

#Returns the ascii value of the char in the string at position POS
Function ASCinString($aString, $pos = 0) {

    $array = [char[]]"$($aString)" | %{ [int]$_ }
    return $array[$pos]

}

# Create a new password, uses 2 words - a noun and verb array of words. Will give you the results where the character lengths matches X
# Ensuring that one word in the password has a letter in it and the other word a special character
Function CreatePassword($length = 10) {
	
	# Create a new password where the words are a certain length of characters.
	# part1 is a pulled from the noun word array.
	# part2 is a pulled from the verb word array.
	# The first char in part1 or part2 may be capatalized.
	# Certain chars of part1 or part2 maybe replaced with new characters.
	# Adding part1 + part2 + random 2 dig number (>9 < 100) to form the password.

    $find =@('ss','d','t','ee','a','i','j','J','z','e')
    $replace =@('$s',')','+','3e','@','!',']',']','2','3')
    
    $part1 = @()  # New array of Nouns 
    $part2 = @()  # New Array of Verbs
	
	# Each password has a number or special character at the end plus a 2 digit number - full password length is (n +1)*2 chars.
    [int]$lgth = (($length-2) /2)

    # create a new nouns array list from words that are only a certain length
    foreach($name in $nouns) {
        if ($name.length -eq $lgth) {
            $part1 += $name
        }
    }
	
	# Create a new verbs array list from words that are only a certain length
    foreach($name in $verbs) {
        if ($name.length -eq $lgth) {
            $part2 += $name
        }
    }

	# Create the 2 part names of the password (name1 , name2) and update one of them with a special and the other a capital character
    do {

        $rnd1 = get-random -minimum 1 -maximum $part1.length
        $rnd2 = get-random -minimum 1 -maximum $part2.length
        $cap = get-random -minimum 1 -maximum 10
        $finalnums = get-random -minimum 10 -maximum 99
        $special = $false
        $number = $false
        $caps = $false
        
		# Select each word from the verb and noun arrays to produce name1 and name2
        $name1= $part1[$rnd1 -1]
        $name2= $part2[$rnd2 -1]
         
        if ($cap -gt 4) {
			
			# Set the first character of name1 to upper, the rest to lower, then check if the first char was changed to an upper character
            $name1 = $name1.substring(0,1).toupper()+$name1.substring(1).tolower() 
            $ascii = ASCinString($name1, 0) # Get the ascii value of the first char
            if ($ascii -ge 65 -AND $ascii -le 90) { $caps = $true } #Set the caps flag to TRUE if the char is now an uppercase ascii 

        } else {
			
			# Set the first character of name2 to upper, the rest to lower, then check if the first char was changed to an upper character
            $name2 = $name2.substring(0,1).toupper()+$name2.substring(1).tolower()
			$ascii = ASCinString($name2, 0) # Get the ascii value of the first char
            if ($ascii -ge 65 -AND $ascii -le 90) { $caps = $true } #Set the caps flag to TRUE if the char is now an uppercase ascii 
			
        }
		
		# Modify one character in each word with a special character, loop until all characters have been checked.
        $i = 0
        do {
			
            $f =$find[$i]
            $r = $replace[$i]
			
            if ($name1 -clike "*$($f)*") {
				
                $regex = [regex]"$($f)"
                [string]$name1 = $regex.Replace($name1, "$($r)", 1)
                $special = $true
				
            } elseif ($name2 -clike "*$($f)*") {
				
                $regex = [regex]"$($f)"
                [string]$name2 = $regex.Replace($name2, "$($r)", 1)
                $special = $true
				
            }

            $i = $i +1
			
        } until ($special -eq $true -OR $i -ge $find.length)

		# Check caps again as the words could not have had the caps letter added previously
		# modify name1, then if it cant, name2 with a uppercase character. (Only some chars will actually change to uppercase)
        if ($caps -eq $false ) {
			
            $ascii = ASCinString($name1, 0) # Get the ascii value of the first char of name1

            if ($ascii -ge 97 -AND $ascii -le 122) { 

                $name1 = $name1.substring(0, 1).toupper()+$name1.substring(1)
                $caps = $true

            } else {
				
				# Now try modifying name2 as name1 failed.
                $ascii = ASCinString($name2, 0) # Get the ascii value of the first char of name2

                if ($ascii -ge 97 -AND $ascii -le 122) { 

                    $name2 = $name2.substring(0,1).toupper()+$name2.substring(1)
                    $caps = $true

                }
            }
        }

    } until ($special -eq $true -AND $caps -eq $true) #Repeat process if the password does not have a special + caps modification
        
	# Create the new password and return it.
    return "$($name1)$($name2)$($finalnums)"

}

# Return a formatted string containing the value plus a size descriptor 
Function Format-Size()
{
    Param ([long]$size)

        If     ($size -gt 1TB) {[string]::Format("{0:0.00}Tb", $size / 1TB)}
        ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00}Gb", $size / 1GB)}
        ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00}Mb", $size / 1MB)}
        ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00}Kb", $size / 1KB)}
        ElseIf ($size -gt 0)   {[string]::Format("{0:0.00}Bytes", $size)}
        Else                   {"0"}
}

# Return a formatted string with the current two digit Hour + Minute. EG 05:59
Function Format-HM()
{
    Return [string]"$(Get-Date -Format("HH:mm"))"
}

# Calculate the Y Location in Pixels of a Form Field to position it to a particular row on the form. 
# This way if the fontsize is incresed, then the fields will be aligned vertically.
# Line is the row down the form, Height is the height in pxels of the current font.
# Use the -Adust to add extra pixels to the value, +Adust to subtract extra Pixels. (EG: ComboBoxes have 2 extra pixels high)
Function GetYLoc($height, $Line = 1, $adjust = 0)
{
    $result = [int](([int]$height * [int]$Line) - ([int]$height -1) +12) - [int]$adjust
    return $result
}

# Calculate the X Location in Pixels of a Form field to position it to a particular column on the form. 
# This way if the fontsize is incresed, then the fields will be aligned horizontally?
# Column on the form, Width is the width in pxels of the current font Character. (Use the X Char as its the widest one)
# Use the -Adust to add extra pixels to the value, +Adust to subtract extra Pixels. 
Function GetXLoc($width, $Column = 1, $adjust = 0)
{
    $result = [int](([int]$width * [int]$Column) - ([int]$width -1) +12) - [int]$adjust
    return $result
}

# Convert a string to a special title case string.
# Some parts of a name needs to be re-converted to ensure they look correct in english.
Function PropperTitleCase($sometext) 
{
  $someText = (Get-Culture).TextInfo.ToTitleCase($someText);
  $someText =$someText.Replace("Macd","MacD");
  $someText =$someText.Replace("Vanh","VanH");
  $someText =$someText.Replace("Van Der ","van der ");
  $someText =$someText.Replace("De La ","de la ");
  $someText =$someText.Replace("Mcc","McC");
  $someText =$someText.Replace("O'f","O'F");
  $someText =$someText.Replace("O'k","O'K");
  $someText =$someText.Replace("O'd","O'D");
  $someText =$someText.Replace("O'b","O'B");
  $someText =$someText.Replace("O'c","O'C");
  $someText =$someText.Replace("-ann ","-Ann ");
  $someText =$someText.Replace("De's","De'S");
  $someText =$someText.Replace("De-l","De-L");
  return $someText;
}

# Add a user to an Active Directory Group
# Use Invoke-Command and pass argument values to the CScriptBlock - ( runs as a separate thread )
Function AddUserToGroup ($group, $dName, $Server) 
{
	
	$ADusername = "powershell.admin"
    $ADpassword = "uglyK!te72"
    $ADsecurePassword = ConvertTo-SecureString $ADpassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($ADusername, $ADsecurePassword)
	
	$scriptBlockGroup = {
        param(
            [string]$groupName,
			[string]$distinguishedName
        )
            
        $check = $null
        try {
            $check = (Get-ADGroup -LDAPFilter "(SAMAccountName=$groupName)" -ErrorAction SilentlyContinue)
        } catch {
            return $_ #Return Failure Message
        }

		$IfInExistingGroup = $null
        if($check) {   #The group does exist - so lets now check if the user is in it?
		    try	{
    	        $IfInExistingGroup = (Get-ADGroupMember -Identity $groupName | where-object { $_.distinguishedName -eq $distinguishedName} -ErrorAction SilentlyContinue)
			} catch {
                return $_  #Return Failure Message
            }
		} else {
            return "The AD Group $group does not exist." # - FAILED
        }
            
		if($IfInExistingGroup) {
			return $True 	#User already a member of the Group - OK
		} else { # Lets try and add the user to the group.
				
			try {
				$AddToGroup = (Add-ADGroupMember -Identity $groupName -Members $distinguishedName -ErrorAction SilentlyContinue)
            } catch {
                return $_   #Return Failure Message
			}
			return $True # User Added to Group - OK
		}
    }

	$result = (Invoke-Command -ComputerName $Server -Credential $credential -ScriptBlock $ScriptBlockGroup -ArgumentList ($group, $dName ))
	
	#True = Succeded - OK
	#False = Not succeded - OK
	#Any other result - FAILED
	
	return $result
}

# Install powershell Tools and Libraries ****************************************************************************************************************************************

$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
$CheckAD = (Get-Module -Name ActiveDirectory -ListAvailable)
if(-not $CheckAD -AND $IsAdmin -eq $False) {
	Start-Process powershell.exe "-NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit
}

if(-not $CheckAD -AND $IsAdmin -eq $True) {
	# Remote Server Administration Tools (RSAT) - Used to create a new user in Active Directory.
	# This tool is designed to help administrators manage and maintain the servers from a remote location
	# Add-WindowsCapability –online –Name “Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0”

	$install = (Get-WindowsCapability -Online | Where-Object {$_.Name -like "RSAT.Active*" -AND $_.State -eq "NotPresent"})

	# Install the RSAT items that meet the filter:
	foreach ($item in $install) {
		try {

			Add-WindowsCapability -Online -Name $item.name

		} catch [System.Exception] {
			Remove_All_Controls($formLoading)
			[void]$formLoading.Close()
			[void]$formLoading.Dispose()
			Write-Warning -Message $_.Exception.Message
			[System.Windows.MessageBox]::Show($_.Exception.Message,"Installing RSAT",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
            Get-Variable | Remove-Variable -ErrorAction SilentlyContinue
			exit
		}
	}
}

if(-not $CheckAD -AND $IsAdmin -eq $False)
{
	Remove_All_Controls($formLoading)
	[void]$formLoading.Close()
	[void]$formLoading.Dispose()
	[System.Windows.MessageBox]::Show("Administrator Login Required to Install RSAT Tools","RSAT",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    Get-Variable | Remove-Variable -ErrorAction SilentlyContinue
    exit
}

# Remove all the Contol Items from a Form. EG: Button, ListBox, ComboBox, Label...
Function Remove_All_Controls($aForm)
{
    $Indexes = @()
    Foreach ($control in $aForm.Controls)
	{
        $Indexes += $control.TabIndex
    }

    Foreach ($index in $indexes | Sort-Object -Descending)
	{
        $aForm.Controls.RemoveAt($index)
    }
}

function Dispose-All-Variables {
    Get-Variable  |
        Where-Object {
            $_.Value -is [System.IDisposable]
        } |
        Foreach-Object {
            $_.Value.Dispose() | Remove-Variable -Force
        }
}

#Ensure Powershell uses TLS when connecting over the network.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Install-ModuleIfNotInstalled 'MSOnline' '1.1.183.66'
Install-ModuleIfNotInstalled 'ExchangeOnlineManagement' '3.2.0'
Install-ModuleIfNotInstalled 'Microsoft.Graph.Users' '2.10.0'
Install-ModuleIfNotInstalled 'Microsoft.Graph.Groups' '2.10.0'
Install-ModuleIfNotInstalled 'Microsoft.Graph.Identity.DirectoryManagement' '2.10.0'
Install-ModuleIfNotInstalled 'Microsoft.Online.SharePoint.PowerShell' '16.0.24211.12000'
Install-ModuleIfNotInstalled 'AzureAD' '2.0.2.180'

#Install-ModuleIfNotInstalled 'Az' '10.0.0'

Import-Module 'ActiveDirectory' -WarningAction SilentlyContinue
Import-Module 'Microsoft.Graph.Users' -WarningAction SilentlyContinue
Import-Module 'Microsoft.Graph.Groups' -WarningAction SilentlyContinue
Import-Module 'Microsoft.Graph.Identity.DirectoryManagement' -WarningAction SilentlyContinue
Import-Module 'ExchangeOnlineManagement' -WarningAction SilentlyContinue
if($Host.Version -gt 5.2) {
	Import-Module 'Microsoft.Online.SharePoint.PowerShell' -WarningAction SilentlyContinue -UseWindowsPowershell
} else {
	Import-Module 'Microsoft.Online.SharePoint.PowerShell' -WarningAction SilentlyContinue
}
Import-Module 'AzureAD'

# Create JobTrackers. This is run from a timer funtion , start tasks in background, and process results once finished.
# The user can start entering in things on the form and the listboxes etc will become enabled with the results in a short time.
$JobTrackerList = New-Object System.Collections.ArrayList
$timerJobTracker = New-Object System.Windows.Forms.Timer
$timerJobTracker.Interval = 900
$timerJobTracker.Enabled = $False

function Add-JobTracker {
    Param(
        [ValidateNotNull()]
        [Parameter(Mandatory=$true)]
        [string]$Name, 
        [ValidateNotNull()]
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$JobScript,
        $ArgumentList = $null,
        [ScriptBlock]$CompletedScript,
        [ScriptBlock]$UpdateScript
    )
    
    #Start the Job
    $job = Start-Job -Name $Name -ScriptBlock $JobScript -ArgumentList $ArgumentList
    
    if($job -ne $null) {

        #Create a Custom Object to keep track of the Job & Script Blocks
        $psObject = New-Object System.Management.Automation.PSObject
        
        Add-Member -InputObject $psObject -MemberType 'NoteProperty' -Name Job  -Value $job
        Add-Member -InputObject $psObject -MemberType 'NoteProperty' -Name CompleteScript  -Value $CompletedScript
        Add-Member -InputObject $psObject -MemberType 'NoteProperty' -Name UpdateScript  -Value $UpdateScript
        
        [void]$JobTrackerList.Add($psObject)    
        
        #Start the Timer
        if(-not $timerJobTracker.Enabled) { $timerJobTracker.Start() }

    } elseif($CompletedScript -ne $null) {

        #Failed
        Invoke-Command -ScriptBlock $CompletedScript -ArgumentList $null
    }

}

function Update-JobTracker
{
    
    $timerJobTracker.Stop() #Freeze the Timer
    
    for($index =0; $index -lt $JobTrackerList.Count; $index++)
    {
        $psObject = $JobTrackerList[$index]
        
        if($psObject -ne $null) {
            
            if($psObject.Job -ne $null) {
                
                if($psObject.Job.State -ne "Running") {                
                    
                    #Call the Complete Script Block
                    if($psObject.CompleteScript -ne $null) {
                        #$results = Receive-Job -Job $psObject.Job
                        Invoke-Command -ScriptBlock $psObject.CompleteScript -ArgumentList $psObject.Job
                    }
                    
                    $JobTrackerList.RemoveAt($index)
                    Remove-Job -Job $psObject.Job
                    $index-- #Step back so we don't skip a job

                } elseif($psObject.UpdateScript -ne $null) {

                    #Call the Update Script Block
                    Invoke-Command -ScriptBlock $psObject.UpdateScript -ArgumentList $psObject.Job

                }
            }
        } else {
            
            $JobTrackerList.RemoveAt($index)
            $index-- #Step back so we don't skip a job

        }
    }
    
    if($JobTrackerList.Count -gt 0) {
        $timerJobTracker.Start()#Resume the timer    
    }    
}

$timerJobTracker_Tick={
    Update-JobTracker
}

function Stop-JobTracker
{
   $timerJobTracker.Stop()
   
   #Remove all the jobs
   while($JobTrackerList.Count-gt 0) {

       $job = $JobTrackerList[0].Job
       $JobTrackerList.RemoveAt(0)
       Stop-Job $job
       Remove-Job $job

   }
}

$script:isMatch  = $false

$domname = ""

try {
	$domname = (Get-ADDomain).NetBIOSName
	$domname = (Get-Culture).TextInfo.ToTitleCase($domname)
} catch { }

$domserver = ""

try {
	$domserver = (Get-ADDomainController).Hostname
	$domserver = $domserver.toUpper()
} catch { }
	
$sendmail = $SendEmailAddress
$lcsendmail = $SendEmailAddress
try {
      $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
      if($searcher) {
            $checksendmail = $searcher.FindOne().Properties.mail
            if($checksendmail) {
                $sendmail = $checksendmail
                $lcsendmail = $sendmail.toLower()
            }
      }
    } catch { }

$domain = ""
$domaindn = ""
$upnDN = ""

# Setup all the Timers
[void]$timerJobTracker.add_Tick($timerJobTracker_Tick)

$main_form = New-Object System.Windows.Forms.Form

$OnLoadForm_StateCorrection= { $main_form.WindowState = $InitialFormWindowState }

$InitialFormWindowState = $main_form.WindowState 
[void]$main_form.add_Load($OnLoadForm_StateCorrection) 

$main_form.Text ="Create a new user for $domname - $domserver"
$main_form.AutoScaleMode = 'Font'
$main_form.Width = 700
$main_form.Height = 544
$main_form.AutoSize = $true
$main_form.TopMost = $true
$main_form.MinimizeBox = $false
$main_form.MaximizeBox = $false
$main_form.icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

$main_form.BackColor = $BackColor
$main_form.ForeColor = $ForeColor

$main_form.FormBorderStyle = 'FixedSingle'  #[FormBorderStyle]::None
$main_form.StartPosition = "CenterScreen"
$main_form.ShowInTaskbar = $true

$main_form_Click = {
    if ($TextBox_User.Text.Length -eq 0 ) {

        [void]$TextBox_User.Focus();

    }
}

[void]$main_form.Add_Click( $main_form_Click )

$main_form_Loading_Shown = {
    $this.Activate();
}
  # Activate the form after it has been painted
[void]$main_form.Add_Shown( $main_form_Loading_Shown )
  
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 

[void]$main_form.SuspendLayout()

$firstLabel  = "User Name"
$aSize = [System.Windows.Forms.TextRenderer]::MeasureText("X", $Form_Font)
[int]$getWidth = $aSize.width + 2 
[int]$getHeight = $aSize.height + 6

$Label_UserName = New-Object System.Windows.Forms.Label
$Label_UserName.Text = $firstLabel
$posY = GetYLoc $getHeight 1
$posX = GetXLoc $getWidth 1
$Label_UserName.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_UserName.AutoSize = $true
$Label_UserName.font = $Form_Font
[void]$main_form.Controls.Add($Label_UserName)

$Label_At = New-Object System.Windows.Forms.Label
$Label_At.Text = $firstLabel
$posX = GetXLoc $getWidth 17 6
$Label_At.Location  = New-Object System.Drawing.Point($posX, $posY)  
$Label_At.AutoSize = $true
$Label_At.font = $Form_Font
$Label_At.text = "@"
[void]$main_form.Controls.Add($Label_At)

$posY = GetYLoc $getHeight 2
$posX = GetXLoc $getWidth 1
$Label_Password = New-Object System.Windows.Forms.Label
$Label_Password.Text = "Password"
$Label_Password.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_Password.AutoSize = $true
$Label_Password.font = $Form_Font
[void]$main_form.Controls.Add($Label_Password)

$posY = GetYLoc $getHeight 2 -3
$posX = GetXLoc $getWidth 17 6
$Label_PasswordChar = New-Object System.Windows.Forms.Label
$Label_PasswordChar.Text = "(Characters)"
$Label_PasswordChar.Location  = New-Object System.Drawing.Point($posX, $posY)  
$Label_PasswordChar.AutoSize = $true
$Label_PasswordChar.font = $Form_Font
[void]$main_form.Controls.Add($Label_PasswordChar)


$Label_FirstName = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 3 -11
$posX = GetXLoc $getWidth 1
$Label_FirstName.Text = "First Name"
$Label_FirstName.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_FirstName.AutoSize = $true
$Label_FirstName.font = $Form_Font
[void]$main_form.Controls.Add($Label_FirstName)

$Label_LastName = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 3 -11
$posX = GetXLoc $getWidth 19
$Label_LastName.Text = "Last Name"
$Label_LastName.Location  = New-Object System.Drawing.Point($posX, $posY)   
$Label_LastName.AutoSize = $true
$Label_LastName.font = $Form_Font
[void]$main_form.Controls.Add($Label_LastName)

$Label_DisplayName = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 4 -11
$posX = GetXLoc $getWidth 1
$Label_DisplayName.Text = "Display Name"
$Label_DisplayName.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_DisplayName.AutoSize = $true
$Label_DisplayName.font = $Form_Font
[void]$main_form.Controls.Add($Label_DisplayName)

$Label_Address = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 14
$posX = GetXLoc $getWidth 1
$Label_Address.Text = "Address"
$Label_Address.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_Address.AutoSize = $true
$Label_Address.font = $Form_Font
$Label_Address.Visible = $false
[void]$main_form.Controls.Add($Label_Address)

$Label_City = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 14
$posX = GetXLoc $getWidth 21
$Label_City.Text = "City"
$Label_City.Location  = New-Object System.Drawing.Point($posX, $posY)  
$Label_City.AutoSize = $true
$Label_City.font = $Form_Font
$Label_City.Visible = $false
[void]$main_form.Controls.Add($Label_City)

$Label_State = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 1
$Label_State.Text = "State"
$Label_State.Location  = New-Object System.Drawing.Point($posX,$posY)
$Label_State.AutoSize = $true
$Label_State.font = $Form_Font
$Label_State.Visible = $false
[void]$main_form.Controls.Add($Label_State)

$Label_PostCode = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 12
$Label_PostCode.Text = "Post Code"
$Label_PostCode.Location  = New-Object System.Drawing.Point($posX, $posY)  
$Label_PostCode.AutoSize = $true
$Label_PostCode.font = $Form_Font
$Label_PostCode.Visible = $false
[void]$main_form.Controls.Add($Label_PostCode)

$Label_Country = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 20
$Label_Country.Text = "Country"
$Label_Country.Location  = New-Object System.Drawing.Point($posX, $posY)  
$Label_Country.AutoSize = $true
$Label_Country.font = $Form_Font
$Label_Country.Visible = $false
[void]$main_form.Controls.Add($Label_Country)

$Label_BasedOn = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 4 -11
$posX = GetXLoc $getWidth 19 -2
$Label_BasedOn.Text = "Setup Like"
$Label_BasedOn.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_BasedOn.AutoSize = $true
$Label_BasedOn.font = $Form_Font
[void]$main_form.Controls.Add($Label_BasedOn)


$Label_Manager = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 6
$posX = GetXLoc $getWidth 19 -7
$Label_Manager.Text = "Manager"
$Label_Manager.Location  = New-Object System.Drawing.Point($posX, $posY) 
$Label_Manager.AutoSize = $true
$Label_Manager.font = $Form_Font
[void]$main_form.Controls.Add($Label_Manager)

$Label_Department = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 6
$posX = GetXLoc $getWidth 1
$Label_Department.Text = "Department"
$Label_Department.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_Department.AutoSize = $true
$Label_Department.font = $Form_Font
[void]$main_form.Controls.Add($Label_Department)

$TextBox_User = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 1
$posX = GetXLoc $getWidth 6 -1
#$posX = 110
$TextBox_User.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_User.Width = 200
$TextBox_User.Name ="user"
$TextBox_User.ForeColor = $ForeColor
$TextBox_User.BackColor= $TextBoxBackColor
$TextBox_User.Font = $TextBoxFont
$TextBox_User.TabIndex = 1
$TextBox_User.MaxLength = 64
Set-Tooltip $TextBox_User "The users name part of the email address. EG: Firstname.Lastname"
[void]$main_form.Controls.Add($TextBox_User)

$TextBox_User_KeyDown = {
    
	if ($_.KeyCode.toString() -in $invalidKeys -or $_.KeyCode -eq 'space') {
       $_.SuppressKeyPress = $True;
    }

}
[void]$TextBox_User.Add_KeyDown($TextBox_User_KeyDown)


$TextBox_User_TextChanged = {
	
    $sel = $TextBox_User.SelectionStart;
    $TextBox_User.Text = $TextBox_User.Text.ToLower();
    $TextBox_User.Text = $TextBox_User.Text.TrimEnd();
    $splitname = $TextBox_User.Text -split "\.";
	
    if($splitname.length -eq 2) {
       if(-not $splitname[1].Contains("@") ) {
           $TextBox_LastName.Text = (Get-Culture).TextInfo.ToTitleCase($splitname[1]);
           $TextBox_DisplayName.Text = $TextBox_FirstName.Text +" " + $TextBox_LastName.Text;
       }
       if(-not $splitname[0].Contains("@") ) {
           $TextBox_FirstName.Text = (Get-Culture).TextInfo.ToTitleCase($splitname[0]);
           $TextBox_DisplayName.Text = $TextBox_FirstName.Text +" " + $TextBox_LastName.Text;
       } 
       if ($splitname[1].length -gt 0) {
            $LabelTick.Text = [Char]8730;
            $LabelTick.forecolor = [System.Drawing.Color]:: Green;
       } else {
            $LabelTick.Text = "";
       }
       if($global:AdUsers) {
            $Checkname = $global:AdUsers | Where-Object {$_.SamAccountName -like "$($TextBox_User.Text)"};
            if($Checkname) {
                $LabelTick.Text = "X";
                $LabelTick.forecolor = [System.Drawing.Color]:: Red; 
            }
       }
       
    } else {
       if(-not $splitname[0].Contains("@") ) {
           $TextBox_FirstName.Text = (Get-Culture).TextInfo.ToTitleCase($splitname[0]);
       }
       $LabelTick.Text = "";
    }
    $findname = $TextBox_User.Text -split "\@";
    if ($findname.length -eq 2) {
       if ($findname[1] -in ($global:UPNS)) {
            $ComboBox_DomainName.Text = $findname[1];
       }
       $TextBox_User.Text = $findname[0];
       $LabelTick.Text = [Char]8730;
       $LabelTick.forecolor = [System.Drawing.Color]:: Green;
	   if($ComboBox_Department.Enabled -eq $true) {
		   [void]$ComboBox_Department.Focus();
	   }
    }
    if ($sel -lt $TextBox_User.Text.Length) {
        $TextBox_User.Select($Sel,0);
    } else {
        $TextBox_User.Select($TextBox_User.Text.Length,0);
    }
}


[void]$TextBox_User.Add_TextChanged($TextBox_User_TextChanged)

$LabelTick = New-Object System.Windows.Forms.Label
$LabelTick.Text = ""
$posY = GetYLoc $getHeight 1 4
$posX = GetXLoc $getWidth 25 8
$LabelTick.Location  = New-Object System.Drawing.Point($posX, $posY)  
$LabelTick.AutoSize = $true
$LabelTick.Visible = $True
$LabelTick.ForeColor = [System.Drawing.Color]::Green
$LabelTick.font = $Tick_Font
[void]$main_form.Controls.Add($LabelTick)

$LabelEmail = New-Object System.Windows.Forms.Label
$LabelEmail.Text = "(Email Address)"
$posY = GetYLoc $getHeight 1 -3
$posX = GetXLoc $getWidth 25 -11
$LabelEmail.Location  = New-Object System.Drawing.Point($posX, $posY) 
$LabelEmail.AutoSize = $true
$LabelEmail.Visible = $True
$LabelEmail.ForeColor = $ForeColor
$LabelEmail.font = $Form_Font
[void]$main_form.Controls.Add($LabelEmail)


$TextBox_Password = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 2
$posX = GetXLoc $getWidth 6 -1
$TextBox_Password.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_Password.Width = 155
$TextBox_Password.Name ="password"
$TextBox_Password.MaxLength = 64
$TextBox_Password.BackColor= $TextBoxBackColor
$TextBox_Password.ForeColor = $ForeColor
$TextBox_Password.Font = $TextBoxFont
$TextBox_Password.TabIndex = 3

Set-Tooltip $TextBox_Password "DoubleClick to Copy Username + Password to Clipboard."
[void]$main_form.Controls.Add($TextBox_Password)

$TextBox_Password_KeyDown = {
	
    if ($_.KeyCode -eq 'Space') {
		
       $_.SuppressKeyPress = $True;
	   
    }
	
}

$TextBox_Password_DoubleClick = {
	
	if ($TextBox_User.Text) {
		$password = "$($TextBox_User.Text)".TrimEnd()+"@$($ComboBox_DomainName.Text)   $($TextBox_Password.Text)".TrimEnd();
	} else {
		$password = "$($TextBox_Password.Text)".TrimEnd();
	}
	
    if (![string]::IsNullOrWhiteSpace($password)) {
		
        [System.Windows.Forms.Clipboard]::SetText($password);
		
    }
	
}

[void]$TextBox_Password.Add_KeyDown( $TextBox_Password_KeyDown )
[void]$TextBox_Password.Add_DoubleClick( $TextBox_Password_DoubleClick )

$TextBox_FirstName = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 3 -11
$posX = GetXLoc $getWidth 6 -1
$TextBox_FirstName.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_FirstName.Width = 130
$TextBox_FirstName.Name ="firstname"
Set-Tooltip $TextBox_FirstName "This field can be populated automatically when entering the Email Address."
$TextBox_FirstName.Font = $TextBoxFont
$TextBox_FirstName.BackColor= $TextBoxBackColor
$TextBox_FirstName.ForeColor = $ForeColor
$TextBox_FirstName.MaxLength = 64
$TextBox_FirstName.TabIndex = 5
[void]$main_form.Controls.Add($TextBox_FirstName)

$TextBox_FirstName_Enter = {
	
    if($TextBox_Password.Text -eq "") {
		
        $TextBox_Password.Text = CreatePassword $ComboBox_PasswordLength.Text;
		
    }
	
}

$TextBox_FirstName_TextChanged = {
  
  $sel = $TextBox_FirstName.SelectionStart;
  $TextBox_FirstName.Text = (Get-Culture).TextInfo.ToTitleCase($TextBox_FirstName.Text);
  
  if ($sel -lt $TextBox_FirstName.Text.Length) {
	  
	  $TextBox_FirstName.Select($sel,0);
	  
  } else {
	  
	$TextBox_FirstName.Select($TextBox_FirstName.Text.Length,0);
	$TextBox_FirstName.ScrollToCaret();
	
  }

}

[void]$TextBox_FirstName.Add_Enter( $TextBox_FirstName_Enter )
[void]$TextBox_FirstName.Add_TextChanged( $TextBox_FirstName_TextChanged )

$TextBox_LastName = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 3 -11
$posX = GetXLoc $getWidth 23
$TextBox_LastName.Location = New-Object System.Drawing.Point($posX, $posY)   
$TextBox_LastName.Width = 130
$TextBox_LastName.Name ="lastname"
$TextBox_LastName.BackColor= $TextBoxBackColor
$TextBox_LastName.ForeColor = $ForeColor
$TextBox_LastName.Font = $TextBoxFont
$TextBox_LastName.TabIndex = 6
Set-Tooltip $TextBox_LastName "This field can be populated automatically when entering the Email Address."
[void]$main_form.Controls.Add($TextBox_LastName)

$TextBox_LastName_TextChanged = {
	
  $sel = $TextBox_LastName.SelectionStart;
  $TextBox_LastName.Text = PropperTitleCase($TextBox_LastName.Text);
  #$TextBox_LastName.SelectionStart = $TextBox_LastName.Text.Length;
  
    if ($sel -lt $TextBox_LastName.Text.Length) {
	  
	  $TextBox_LastName.Select($sel,0);
	  
  } else {
	  
		$TextBox_LastName.Select($TextBox_LastName.Text.Length,0);
		$TextBox_LastName.ScrollToCaret();
  }

}
$TextBox_LastName_Enter = {
    if($TextBox_Password.Text -eq "") {
        $TextBox_Password.Text = CreatePassword $ComboBox_PasswordLength.Text;
    }
    [System.Windows.Forms.Application]::DoEvents();
}
[void]$TextBox_LastName.Add_TextChanged( $TextBox_LastName_TextChanged )
[void]$TextBox_LastName.Add_Enter( $TextBox_LastName_Enter )

$TextBox_DisplayName = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 4 -11
$posX = GetXLoc $getWidth 6 -1
$TextBox_DisplayName.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_DisplayName.Width = 220
$TextBox_DisplayName.Name ="displayname"
$TextBox_DisplayName.BackColor= $TextBoxBackColor
$TextBox_DisplayName.ForeColor = $ForeColor
$TextBox_DisplayName.Font = $TextBoxFont
$TextBox_DisplayName.MaxLength = 256
$TextBox_DisplayName.TabIndex = 7
Set-Tooltip $TextBox_DisplayName "This field will be automatically created when entering the email address."
[void]$main_form.Controls.Add($TextBox_DisplayName)

$TextBox_DisplayName_LostFocus = {
	
  $TextBox_DisplayName.Text = PropperTitleCase($TextBox_DisplayName.Text);
  
}
[void] $TextBox_DisplayName.Add_LostFocus( $TextBox_DisplayName_LostFocus)

$TextBox_Address = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 14
$posX = GetXLoc $getWidth 6 -1
$TextBox_Address.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_Address.Width = 270
$TextBox_Address.Name ="Address"
$TextBox_Address.BackColor= $TextBoxBackColor
$TextBox_Address.ForeColor = $ForeColor
$TextBox_Address.Font = $TextBoxFont
$TextBox_Address.Text = $AUStreetAddress
$TextBox_Address.Visible = $false
Set-Tooltip $TextBox_Address "The physical office location for the user."
[void]$main_form.Controls.Add($TextBox_Address)

$TextBox_City = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 14
$posX = GetXLoc $getWidth 23
$TextBox_City.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_City.Width = 100
$TextBox_City.Name = "City"
$TextBox_City.BackColor= $TextBoxBackColor
$TextBox_City.ForeColor = $ForeColor
$TextBox_City.Font = $TextBoxFont
$TextBox_City.Text = $AUCity 
$TextBox_City.Visible = $false
$TextBox_City.MaxLength = 128
Set-Tooltip $TextBox_City "The phyisical office location for the user."
[void]$main_form.Controls.Add($TextBox_City)

$TextBox_City_LostFocus = {
	
  $TextBox_City.Text = PropperTitleCase($TextBox_City.Text);
  
}
[void] $TextBox_City.Add_LostFocus( $TextBox_City_LostFocus)

$TextBox_State = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 6 -1
$TextBox_state.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_State.Width = 100
$TextBox_State.Name = "State"
$TextBox_State.BackColor= $TextBoxBackColor
$TextBox_State.ForeColor = $ForeColor
$TextBox_State.Font = $TextBoxFont
$TextBox_State.Text = $AUState
$TextBox_State.Visible = $false
Set-Tooltip $TextBox_State "The physical office state location for the user."
[void]$main_form.Controls.Add($TextBox_State)

$TextBox_State_LostFocus= {
  $TextBox_State.Text = PropperTitleCase($TextBox_State.Text);
  
}
[void] $TextBox_State.Add_LostFocus( $TextBox_State_LostFocus)

$TextBox_PostCode = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 16
$TextBox_PostCode.Location = New-Object System.Drawing.Point($posX, $posY)  
$TextBox_PostCode.Width = 50
$TextBox_PostCode.Name = "PostCode"
$TextBox_PostCode.BackColor= $TextBoxBackColor
$TextBox_PostCode.ForeColor = $ForeColor
$TextBox_PostCode.Font = $TextBoxFont
$TextBox_PostCode.Text = $AUpostalcode
$TextBox_PostCode.Visible = $false
$TextBox_PostCode.MaxLength = 4
Set-Tooltip $TextBox_PostCode "The postal code of the physical location for the user."
[void]$main_form.Controls.Add($TextBox_PostCode)

$ComboBox_Country = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 15
$posX = GetXLoc $getWidth 23 -1
$ComboBox_Country.Location = New-Object System.Drawing.Point($posX, $posY) 
$ComboBox_Country.Width = 99
$ComboBox_Country.Name = "Country"
$ComboBox_Country.BackColor= $TextBoxBackColor
$ComboBox_Country.ForeColor = $ForeColor
$ComboBox_Country.Font = $TextBoxFont
$ComboBox_Country.Text = $AUcountry
$ComboBox_Country.Visible = $false
Set-Tooltip $ComboBox_Country "The country code for the physical location for the user."
$ComboBox_Country.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox_Country.FlatStyle = "Flat"
[void]$main_form.Controls.Add($ComboBox_Country)

[void]$ComboBox_Country.Items.Add($AUcountry)
[void]$ComboBox_Country.Items.Add($SecondCountry)
$ComboBox_Country.SelectedIndex = 0


$ComboBox_BasedOn = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 4 -11
$posX = GetXLoc $getWidth 23 -1
$ComboBox_BasedOn.Width = 220
$ComboBox_BasedOn.Name = "BasedOn"
$ComboBox_BasedOn.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_BasedOn.Location  = New-Object System.Drawing.Point($posX, $posY) 
$ComboBox_BasedOn.BackColor= $TextBoxBackColor
$ComboBox_BasedOn.ForeColor = $ForeColor
$ComboBox_BasedOn.FlatStyle = "Flat"
$ComboBox_BasedOn.Font = $TextBoxFont
Set-Tooltip $ComboBox_BasedOn "Set the users Group Membership to match the selected user from in this list. (Dont select if your not sure?)"
$ComboBox_BasedOn.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox_BasedOn.Enabled = $false
$ComboBox_BasedOn.MaxLength = 64
$ComboBox_BasedOn.TabIndex = 9
[void]$main_form.Controls.Add($ComboBox_BasedOn)

$ComboBox_Manager = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 6
$posX = GetXLoc $getWidth 23 -1
$ComboBox_Manager.Width = 220
$ComboBox_Manager.Name = "Manager"
$ComboBox_Manager.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_Manager.Location  = New-Object System.Drawing.Point($posX, $posY) 
$ComboBox_Manager.BackColor= $TextBoxBackColor
$ComboBox_Manager.ForeColor = $ForeColor
$ComboBox_Manager.FlatStyle = "Flat"
$ComboBox_Manager.Font = $TextBoxFont
Set-Tooltip $ComboBox_Manager "If advised, select the users Manager from this list.`nEmployees can be listed under each Manager in Outlook."
$ComboBox_Manager.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox_Manager.Enabled = $false
$ComboBox_Manager.MaxLength = 64
$ComboBox_Manager.TabIndex = 9
[void]$main_form.Controls.Add($ComboBox_Manager)



$ComboBox_Department = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 6
$posX = GetXLoc $getWidth 6 -2
$ComboBox_Department.Width = 220
$ComboBox_Department.Name = "Department"
$ComboBox_Department.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_Department.Location  = New-Object System.Drawing.Point($posX, $posY) 
$ComboBox_Department.BackColor= $TextBoxBackColor
$ComboBox_Department.ForeColor = $ForeColor
$ComboBox_Department.FlatStyle = "Flat"
$ComboBox_Department.Font = $TextBoxFont
$ComboBox_Department.MaxLength = 64
$ComboBox_Department.TabIndex = 8
Set-Tooltip $ComboBox_Department "If Advised, select a Department from this list, or type in a New One."
$ComboBox_Department.Enabled = $false
[void]$main_form.Controls.Add($ComboBox_Department)

$CheckBox_UseLaptopOrPC = New-Object System.Windows.Forms.RadioButton
$posY = GetYLoc $getHeight 7 -2
$CheckBox_UseLaptopOrPC.Text = "This user is going to be using a Laptop or Desktop PC"
$CheckBox_UseLaptopOrPC.Location  = New-Object System.Drawing.Point($PosX, $posY)
$CheckBox_UseLaptopOrPC.Width = 400
$CheckBox_UseLaptopOrPC.font = $Form_Font
$CheckBox_UseLaptopOrPC.UseVisualStyleBackColor = $True
$CheckBox_UseLaptopOrPC.TabIndex = 10
Set-Tooltip $CheckBox_UseLaptopOrPC "The User will be issued with a Business Premium license."
[void]$main_form.Controls.Add($CheckBox_UseLaptopOrPC)

$CheckBox_UseLaptopOrPC_Click = {
	
   if($CheckBox_UseLaptopOrPC.Checked -eq $False) {

        $CheckBox_SecondOffice.Checked = $False;
        $Label_Title.visible = $False;
		$Label_Address.visible = $False;
		$Label_City.Visible = $False;
		$Label_State.Visible = $False;	
		$Label_PostCode.Visible = $False;
		$Label_Country.Visible = $False;
		$TextBox_Address.visible = $False;
		$TextBox_City.Visible = $False;
		$TextBox_State.Visible = $False;
		$TextBox_PostCode.Visible = $False;
        $ComboBox_Title.visible = $False;
        $Label_Mobile.visible = $False;
        $TextBox_Mobile.visible = $False;
        $Label_Phone.visible = $False;
        $TextBox_Phone.visible = $False;
        $Label_MobileEG.visible = $False;
        $Label_PhoneEG.visible = $False;
		$ComboBox_Country.Visible = $false;

        if($TextBox_User.Text -eq "") {
            [void] $TextBox_User.Focus();
        }
		
        if($TextBox_Password.Text -eq "") {
            [void] $TextBox_Password.Focus();
        }
		
    } else {

        $ComboBox_License.SelectedIndex = 4;
        $Label_Title.visible = $True;
        $ComboBox_Title.visible = $True;
		$Label_Address.visible = $True;
		$Label_City.Visible = $True;
		$Label_State.Visible = $True;
		$Label_PostCode.Visible = $True;
		$Label_Country.Visible = $True;
		$TextBox_Address.visible = $True;
		$TextBox_City.Visible = $True;
		$TextBox_State.Visible = $True;
		$TextBox_PostCode.Visible = $True;
        $Label_Mobile.visible = $True;
        $TextBox_Mobile.visible = $True;
        $Label_Phone.visible = $True;
        $TextBox_Phone.visible = $True;
        $Label_MobileEG.visible = $True;
        $Label_PhoneEG.visible = $True;
		$ComboBox_Country.Visible = $True;

        if(-not $TextBox_Password.Text -eq "") {
            if($ComboBox_Title.Text -eq "") { 
				if($ComboBox_Title.Enabled -eq $True) {
					[void] $ComboBox_Title.Focus();
				}
			}
        } else { 
            [void] $TextBox_User.Focus()
        }
    }
    [System.Windows.Forms.Application]::DoEvents();
}

[void]$CheckBox_UseLaptopOrPC.Add_Click( $CheckBox_UseLaptopOrPC_Click )

$CheckBox_AccessMobileOrBrowser = New-Object System.Windows.Forms.RadioButton
$posY = GetYLoc $getHeight 8 4
$CheckBox_AccessMobileOrBrowser.Text = "This user will only access email via Mobile or Browser"
$CheckBox_AccessMobileOrBrowser.Location  = New-Object System.Drawing.Point($posX, $PosY)
$CheckBox_AccessMobileOrBrowser.Width = 400
$CheckBox_AccessMobileOrBrowser.Checked = $true
$CheckBox_AccessMobileOrBrowser.font = $Form_Font
$CheckBox_AccessMobileOrBrowser.UseVisualStyleBackColor = $True
$CheckBox_AccessMobileOrBrowser.TabIndex = 11
Set-Tooltip $CheckBox_AccessMobileOrBrowser "The user will be issued with an E1 license."
[void]$main_form.Controls.Add($CheckBox_AccessMobileOrBrowser)

$CheckBox_AccessMobileOrBrowser_Click = {

    if($CheckBox_AccessMobileOrBrowser.Checked -eq $True ) {

        $ComboBox_License.SelectedIndex = 0;
        $CheckBox_SecondOffice.Checked = $False;
        $Label_Title.visible = $False;
        $ComboBox_Title.visible = $False;
		$Label_Address.visible = $False;
		$Label_City.Visible = $False;
		$Label_State.Visible = $False;
		$Label_PostCode.Visible = $False;
		$Label_Country.Visible = $False;
		$TextBox_Address.visible = $False;
		$TextBox_City.Visible = $False;
		$TextBox_State.Visible = $False;
		$TextBox_PostCode.Visible = $False;
        $Label_Mobile.visible = $False;
        $TextBox_Mobile.visible = $False;
        $Label_Phone.visible = $False;
        $TextBox_Phone.visible = $False;
        $Label_MobileEG.visible = $False;
        $Label_PhoneEG.visible = $False;
		$ComboBox_Country.Visible = $False;

        if($TextBox_User.Text -eq "") {
            [void] $TextBox_User.Focus();
        }

        if($TextBox_Password.Text -eq "") {
            [void] $TextBox_Password.Focus();
        }

    } else {

        $Label_Title.visible = $True;
        $ComboBox_Title.visible = $True;
		$Label_Address.visible = $True;
		$Label_City.Visible = $True;
		$Label_State.Visible = $True;
		$Label_PostCode.Visible = $True;
		$Label_Country.Visible = $True;
		$TextBox_Address.visible = $True;
		$TextBox_City.Visible = $True;
		$TextBox_State.Visible = $True;
		$TextBox_PostCode.Visible = $True;
        $Label_Mobile.visible = $True;
        $TextBox_Mobile.visible = $True;
        $Label_Phone.visible = $True;
        $TextBox_Phone.visible = $True;
        $Label_MobileEG.visible = $True;
        $Label_PhoneEG.visible = $True;
		$ComboBox_Country.Visible = $True;

        if(-not $TextBox_Password.Text -eq "") {
            if($ComboBox_Title.Text -eq "") {
				if($ComboBox_Title.Enabled -eq $True) {
					[void] $ComboBox_Title.Focus();
				}
			}
        } else {
            [void] $TextBox_User.Focus()
        }
    }

    [System.Windows.Forms.Application]::DoEvents();
}

[void]$CheckBox_AccessMobileOrBrowser.Add_Click( $CheckBox_AccessMobileOrBrowser_Click )

$CheckBox_SecondOffice = New-Object System.Windows.Forms.CheckBox
$posY = GetYLoc $getHeight 9 8
$CheckBox_SecondOffice.Text = "This user is located in the $($SecondCountry) Office"
$CheckBox_SecondOffice.Location  = New-Object System.Drawing.Point($posX, $PosY)
$CheckBox_SecondOffice.Width = 400
$CheckBox_SecondOffice.Checked = $false
$CheckBox_SecondOffice.font = $Form_Font
$CheckBox_SecondOffice.UseVisualStyleBackColor = $True
$CheckBox_SecondOffice.TabIndex = 12
Set-Tooltip $CheckBox_SecondOffice "The user will be issued with a Business Premium license."
[void]$main_form.Controls.Add($CheckBox_SecondOffice)

$CheckBox_SecondOffice_CheckStateChanged = {

    if($CheckBox_SecondOffice.Checked -eq $True ) {

        $ComboBox_License.SelectedIndex = 4;
        $CheckBox_UseLaptopOrPC.Checked = $True;
        $Label_Title.visible = $True;
        $ComboBox_Title.visible = $True;
		$Label_Address.visible = $True;
		$Label_City.Visible = $True;
		$Label_State.Visible = $True;
		$Label_PostCode.Visible = $True;
		$Label_Country.Visible = $True;
		$ComboBox_Country.SelectedIndex = 1
		$TextBox_Address.Text = $SecondStreetAddress;
		$TextBox_City.Text = $SecondCity;
		$TextBox_State.Text = $SecondState;
		$TextBox_PostCode.Text = $Secondpostalcode;
		$TextBox_Address.visible = $True;
		$TextBox_City.Visible = $True;
		$TextBox_State.Visible = $True;
		$TextBox_PostCode.Visible = $True;
        $Label_Mobile.visible = $True;
        $TextBox_Mobile.visible = $True;
        $Label_Phone.visible = $True;
        $TextBox_Phone.visible = $True;
        $Label_MobileEG.visible = $True;
        $Label_PhoneEG.visible = $True;
		$ComboBox_Country.Visible = $True;

        if(-not $TextBox_Password.Text -eq "") {
			
            if($ComboBox_Title.Text -eq "") {
				if($ComboBox_Title.Enabled -eq $True) {
					[void] $ComboBox_Title.Focus();
				}
			}
			
        } else {
            [void] $TextBox_User.Focus()
        }

    } else {

		$TextBox_Address.Text = $AUstreetaddress;
		$TextBox_City.Text = $AUcity;
		$TextBox_State.Text = $AUstate;
		$TextBox_PostCode.Text = $AUpostalcode;
		$ComboBox_Country.SelectedIndex = 0
		
	}
    [System.Windows.Forms.Application]::DoEvents();
}

[void] $CheckBox_SecondOffice.Add_CheckStateChanged( $CheckBox_SecondOffice_CheckStateChanged )

$CheckBox_SendEmail = New-Object System.Windows.Forms.CheckBox
$posY = GetYLoc $getHeight 18 5
$CheckBox_SendEmail.Text = "Email a copy of this information to $lcsendmail"
$CheckBox_SendEmail.font = $Form_Font

#Set the Checkbox to the Centre of the Form. (Add 18px for the checkbox itself)
$newSize = [System.Windows.Forms.TextRenderer]::MeasureText($CheckBox_SendEmail.Text, $CheckBox_SendEmail.font)
[int]$Size = $newSize.Width + 18
[int]$cntre = ($main_form.Width - $Size) / 2
$CheckBox_SendEmail.Location  = New-Object System.Drawing.Point($cntre, $PosY)
$CheckBox_SendEmail.Width = $Size

$CheckBox_SendEmail.UseVisualStyleBackColor = $True
$CheckBox_SendEmail.Checked = $True
Set-Tooltip $CheckBox_SendEmail "If checked, send an email with all the users details."
[void]$main_form.Controls.Add($CheckBox_SendEmail)


$Label_Title = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 11 -11
$posX = GetXLoc $getWidth 1
$Label_Title.Text = "Title"
$Label_Title.Location  = New-Object System.Drawing.Point($PosX, $PosY)
$Label_Title.AutoSize = $true
$Label_Title.Visible = $False
$Label_Title.font = $Form_Font
[void]$main_form.Controls.Add($Label_Title)

$posY = GetYLoc $getHeight 11 -11
$posX = GetXLoc $getWidth 6 -2
$ComboBox_Title = New-Object System.Windows.Forms.ComboBox
$ComboBox_Title.Location = New-Object System.Drawing.Point($posX, $posY)
$ComboBox_Title.Width = 300
$ComboBox_Title.Name = "Title"
$ComboBox_Title.Visible = $False
$ComboBox_Title.BackColor= $TextBoxBackColor
$ComboBox_Title.ForeColor = $ForeColor
$ComboBox_Title.Font = $TextBoxFont
$ComboBox_Title.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_Title.FlatStyle = "Flat"
$ComboBox_Title.MaxLength = 128

Set-Tooltip $ComboBox_Title "Enter a users title, or select from a list of currently used ones.`nThis will appear in the Users Signature."
[void]$main_form.Controls.Add($ComboBox_Title)

$ComboBox_Title_LostFocus = {
  $ComboBox_Title.Text = PropperTitleCase($ComboBox_Title.Text);
  [System.Windows.Forms.Application]::DoEvents();
}

[void] $ComboBox_Title.Add_LostFocus( $ComboBox_Title_LostFocus)

$ComboBox_Title_SelectedIndexChanged = {

  if ( $TextBox_Mobile.text -eq "") {$TextBox_Mobile.Focus()}

}

[void] $ComboBox_Title.Add_SelectedIndexChanged( $ComboBox_Title_SelectedIndexChanged )

$Label_Mobile = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 12 -11
$posX = GetXLoc $getWidth 1
$Label_Mobile.Text = "Mobile"
$Label_Mobile.Location  = New-Object System.Drawing.Point($PosX, $posY)
$Label_Mobile.AutoSize = $true
$Label_Mobile.Visible = $False
$Label_Mobile.font = $Form_Font
[void]$main_form.Controls.Add($Label_Mobile)

$TextBox_Mobile = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 12 -11
$posX = GetXLoc $getWidth 6 -1
$TextBox_Mobile.Location = New-Object System.Drawing.Point($posX, $posY)
$TextBox_Mobile.Width = 100
$TextBox_Mobile.Name = "Mobile"
$TextBox_Mobile.Visible = $False
$TextBox_Mobile.BackColor= $TextBoxBackColor
$TextBox_Mobile.ForeColor = $ForeColor
$TextBox_Mobile.Font = $TextBoxFont
$TextBox_Mobile.MaxLength = 64
Set-Tooltip $TextBox_Mobile "The mobile phone is displayed in the users Email Signature."
[void]$main_form.Controls.Add($TextBox_Mobile)

$TextBox_Mobile_LostFocus= {

	#Format Number only if its an Australian Address
	if($CheckBox_SecondOffice -eq $False) {

		$phone = $TextBox_Mobile.Text.replace(" ","");

		if ($phone.Length -eq 10) { 
			$phone= ([int64]$phone).ToString('0### ### ###');
		}

		$TextBox_Mobile.Text = $phone;
	}

	if($CheckBox_SecondOffice -eq $True) {

		$phone = $TextBox_Mobile.Text.replace(" ","");

		if ($phone.Length -eq 10) { 
			$phone= ([int64]$phone).ToString('0## ### ####');
		}

		$TextBox_Mobile.Text = $phone;
	}
	 
}

[void] $TextBox_Mobile.Add_LostFocus( $TextBox_Mobile_LostFocus)

$Label_MobileEG= New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 12 -15
$posX = GetXLoc $getWidth 12
$Label_MobileEG.Text = "EG: 0412 555 111"
$Label_MobileEG.Location  = New-Object System.Drawing.Point($posX, $posY) 
$Label_MobileEG.AutoSize = $true
$Label_MobileEG.Visible = $False
[void]$main_form.Controls.Add($Label_MobileEG)

$Label_Phone = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 12 -11
$posX = GetXLoc $getWidth 20 -8
$Label_Phone.Text = "Phone"
$Label_Phone.Location  = New-Object System.Drawing.Point($posX, $posY) 
$Label_Phone.AutoSize = $true
$Label_Phone.Visible = $False
$Label_Phone.font = $Form_Font
[void]$main_form.Controls.Add($Label_Phone)

$TextBox_Phone = New-Object System.Windows.Forms.TextBox
$posY = GetYLoc $getHeight 12 -11
$posX = GetXLoc $getWidth 23
$TextBox_Phone.Location = New-Object System.Drawing.Point($posX, $posY)  
$TextBox_Phone.Width = 100
$TextBox_Phone.Name = "Phone"
$TextBox_Phone.Visible = $False
$TextBox_Phone.BackColor= $TextBoxBackColor
$TextBox_Phone.ForeColor = $ForeColor
$TextBox_Phone.Font = $TextBoxFont
$TextBox_Phone.MaxLength = 64
Set-Tooltip $TextBox_Phone "The direct phone is displayed in the users Email Signature."
[void]$main_form.Controls.Add($TextBox_Phone)

$TextBox_Phone_LostFocus = {

	#Format Number only if its an Australian Address
	if($CheckBox_SecondOffice -eq $False) {

		$phone = $TextBox_Phone.Text.replace(" ","");
		if ($phone.Length -eq 10) { $phone= ([int64]$phone).ToString('0# #### ####') }
		$TextBox_Phone.Text = $phone;

	}
}

[void] $TextBox_Phone.Add_LostFocus( $TextBox_Phone_LostFocus )

$Label_PhoneEG = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 12 -15
$posX = GetXLoc $getWidth 29
$Label_PhoneEG.Text = "EG: 07 3555 2233"
$Label_PhoneEG.Location  = New-Object System.Drawing.Point($posX, $posY)
$Label_PhoneEG.AutoSize = $true
$Label_PhoneEG.Visible = $False
[void]$main_form.Controls.Add($Label_PhoneEG)

$Label_License = New-Object System.Windows.Forms.Label
$posY = GetYLoc $getHeight 10
$posX = GetXLoc $getWidth 1
$Label_License.Text = "365 License"
$Label_License.Location  = New-Object System.Drawing.Point($PosX, $posY)
$Label_License.AutoSize = $true
$Label_License.Visible = $true
$Label_License.font = $Form_Font
[void]$main_form.Controls.Add($Label_License)

$ComboBox_License = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 10
$posX = GetXLoc $getWidth 6 -2
$ComboBox_License.Width = 300
$ComboBox_License.Name = "License"
$ComboBox_License.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_License.Location  = New-Object System.Drawing.Point($posX, $posY) 
$ComboBox_License.BackColor= $TextBoxBackColor
$ComboBox_License.ForeColor = $ForeColor
$ComboBox_License.FlatStyle = "Flat"
$ComboBox_License.Font = $TextBoxFont
$ComboBox_License.Text = ""
$ComboBox_License.TabIndex = 13
Set-Tooltip $ComboBox_License "A monthly charge is applicable to each allocated license."
$ComboBox_License.DataBindings.DefaultDataSourceUpdateMode = 0
$ComboBox_License.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

[void]$main_form.Controls.Add($ComboBox_License)
[void] $ComboBox_License.Items.Add("Office 365 E1 (50Gb Online Only)")
[void] $ComboBox_License.Items.Add("Office 365 E3 (100Gb Mailbox, Desktop, Archive)")
[void] $ComboBox_License.Items.Add("Office 365 E5 (100Gb Mailbox, E3 + VOIP)")
[void] $ComboBox_License.Items.Add("Business Essentials (50Gb Online Only)")
[void] $ComboBox_License.Items.Add("Business Premium (50Gb Mailbox, Desktop)")
$ComboBox_License.SelectedIndex = 0

$ComboBox_PasswordLength = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 2
$posX = GetXLoc $getWidth 14 -6
$ComboBox_PasswordLength.Width = 40
$ComboBox_PasswordLength.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_PasswordLength.Location  = New-Object System.Drawing.Point($posX, $posY)  
$ComboBox_PasswordLength.BackColor= $TextBoxBackColor
$ComboBox_PasswordLength.ForeColor = $ForeColor
$ComboBox_PasswordLength.FlatStyle = "Flat"
$ComboBox_PasswordLength.Font = $TextBoxFont
$ComboBox_PasswordLength.Text = "10"
$ComboBox_PasswordLength.TabIndex = 4
Set-Tooltip $ComboBox_PasswordLength "Change the number of characters in the password."
[void]$main_form.Controls.Add($ComboBox_PasswordLength)
[void] $ComboBox_PasswordLength.Items.Add("8")
[void] $ComboBox_PasswordLength.Items.Add("10")
[void] $ComboBox_PasswordLength.Items.Add("12")
[void] $ComboBox_PasswordLength.Items.Add("14")
[void] $ComboBox_PasswordLength.Items.Add("16")
[void] $ComboBox_PasswordLength.Items.Add("18")
[void] $ComboBox_PasswordLength.Items.Add("20")
[void] $ComboBox_PasswordLength.Items.Add("22")
[void] $ComboBox_PasswordLength.Items.Add("24")

$ComboBox_PasswordLength_SelectedIndexChanged = {

    $num = $this.SelectedItem;
    $TextBox_Password.Text = CreatePassword $num;

}

[void] $ComboBox_PasswordLength.Add_SelectedIndexChanged( $ComboBox_PasswordLength_SelectedIndexChanged )

$ComboBox_DomainName = New-Object System.Windows.Forms.ComboBox
$posY = GetYLoc $getHeight 1
$posX = GetXLoc $getWidth 18 6
$ComboBox_DomainName.Width = 130
$ComboBox_DomainName.Cursor = [System.Windows.Forms.Cursors]::Hand
$ComboBox_DomainName.Location  = New-Object System.Drawing.Point($posX, $posY)  
$ComboBox_DomainName.BackColor= $TextBoxBackColor
$ComboBox_DomainName.ForeColor = $ForeColor
$ComboBox_DomainName.FlatStyle = "Flat"
$ComboBox_DomainName.Font = $TextBoxFont
$ComboBox_DomainName.Text = $EmailDomain  #Default Option
$ComboBox_DomainName.Enabled = $false
$ComboBox_DomainName.TabIndex = 2

Set-Tooltip $ComboBox_DomainName "Change the domain for the users email address."
[void]$main_form.Controls.Add($ComboBox_DomainName)

[int]$buttonPosY = [int]$main_form.Height - 72

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Size(390, $buttonPosY)
$okButton.Size = New-Object System.Drawing.Size(120,24)
$okButton.Text = "OK (Create)"
$okButton.FlatAppearance.BorderColor = $ButtonBorderColor
$okButton.FlatAppearance.BorderSize = 2
$okButton.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
$okButton.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
$okButton.FlatStyle = "Flat"
$okButton.ForeColor = $ForeColor
$okButton.BackColor = $ButtonColor
$okButton.font = $ButtonFont

$okButton_Click = {

   if($TextBox_User.Text -eq "") {

        [System.Windows.Forms.MessageBox]::Show("Please enter a 'User Name' to create the new user. `nEG: FirstName.LastName   (@$($EmailDomain))", "ERROR: missing User Name",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
        $TextBox_User.Focus();

    } else {

        if($TextBox_FirstName.Text -eq "") {

            [System.Windows.Forms.MessageBox]::Show("Please enter a 'First Name' (It must not be blank) `nEG: Mark", "ERROR: missing First Name",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            $TextBox_FirstName.Focus();

        } else {

            if($TextBox_LastName.Text -eq "") {

                [System.Windows.Forms.MessageBox]::Show("Please enter a 'Last Name' (It must not be blank) `nEG: Minion", "ERROR: missing Last Name",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
                $TextBox_LastName.Focus();

            } else {

                if($TextBox_Password.Text -eq "") {

                    [System.Windows.Forms.MessageBox]::Show("Please enter a 'Password' (It must not be blank)  `nEG: Bl@ckMark23", "ERROR: missing Password",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
                    $TextBox_Password.Focus();

                } else {
					
						$main_form.Tag = $TextBox_Password.Text;
						Remove_All_Controls($main_form);
						[void]$main_form.Close();
					
                }
            }
        }
    }
   [System.Windows.Forms.Application]::DoEvents();
}

[void]$okButton.Add_Click( $okButton_Click )


$okButton_Paint = {

    $hrgn = $Win32Helpers::CreateRoundRectRgn(0,0,$okButton.Width, $okButton.Height, 3,3);
    $okButton.Region = [Region]::FromHrgn($hrgn);
    [System.Windows.Forms.Application]::DoEvents();

}

[void]$okButton.add_Paint( $okButton_Paint )

# Create the Cancel button.
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Size(120,$buttonPosY)
$cancelButton.Size = New-Object System.Drawing.Size(120,24)
$cancelButton.Text = "Cancel"
$cancelButton.font = $ButtonFont
$cancelButton.FlatAppearance.BorderColor = $ButtonBorderColor
$cancelButton.FlatAppearance.BorderSize = 2
$cancelButton.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
$cancelButton.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
$cancelButton.FlatStyle = "Flat"
$cancelButton.ForeColor = $ForeColor
$cancelButton.BackColor = $ButtonColor

$cancelButton_Click = {

   $main_form.Tag = $null;
   Remove_All_Controls($main_form);
   $main_form.Close();
   [System.Windows.Forms.Application]::DoEvents();
      
}

$cancelButton_Paint = {

    $hrgn = $Win32Helpers::CreateRoundRectRgn(0,0,$cancelButton.Width, $cancelButton.Height, 3,3);
    $cancelButton.Region = [Region]::FromHrgn($hrgn);
    [System.Windows.Forms.Application]::DoEvents();

}

[void]$cancelButton.Add_Click( $cancelButton_Click )
[void]$cancelButton.add_Paint( $cancelButton_Paint )

[void]$main_form.AcceptButton.$okButton
[void]$main_form.CancelButton.$cancelButton
[void]$main_form.Controls.Add($okButton)
[void]$main_form.Controls.Add($cancelButton)

[System.Windows.Forms.Application]::UseWaitCursor = $False

$TextBox_Password.Text = CreatePassword 10
[System.Windows.Forms.Application]::DoEvents()

# Create all the background Jobs that are checked for completion every 1 secs
# The background jobs run as new process and gather data from an active drirectory server and place results into combox boxes.
Add-JobTracker -Name "GetManagerNames" `
    -JobScript {
         get-aduser -SearchBase $using:OUPath -LDAPFilter '(!userAccountControl:1.2.840.113556.1.4.803:=2)' -Properties Name
    }`
    -CompletedScript {
        Param($Job)
        $result = Receive-Job -Job $Job -Keep
        if($result) {
          $array = New-Object System.Collections.ArrayList
          Foreach ($name in $result) {
             $array.Add($name.Name) > $null
          }

          if($ComboBox_Manager) {
			$ComboBox_Manager.Items.clear()
            [void] $ComboBox_Manager.Items.Add("")
            Foreach ($User in ($array | Sort-Object) )
            {
              [void] $ComboBox_Manager.Items.Add($User)
            }
			$ComboBox_Manager.Enabled = $True
            [System.Windows.Forms.Application]::DoEvents()
          }
		  if($ComboBox_BasedOn) {
			$ComboBox_BasedOn.Items.clear()
            [void] $ComboBox_BasedOn.Items.Add("")
            Foreach ($User in ($array | Sort-Object) )
            {
              [void] $ComboBox_BasedOn.Items.Add($User)
            }
			$ComboBox_BasedOn.Enabled = $True
            [System.Windows.Forms.Application]::DoEvents()
          }
        }
    }`
    -UpdateScript {
    }

Add-JobTracker -Name "GetDepartmentNames" `
    -JobScript {
        try {
        (get-aduser -filter "enabled -eq 'true'" -property department).department | Sort-Object -Unique
        }
        catch {}
    }`
    -CompletedScript {
        Param($Job)
           $result = Receive-Job -Job $Job -Keep
           if($result) {
                $array = New-Object System.Collections.ArrayList
                Foreach ($name in $result) {
                    $array.Add($name) > $null
                }
              if($ComboBox_Department) {
				  $ComboBox_Department.Items.clear()
                Foreach ($name in ($array | Sort-Object))
                {
                    [void] $ComboBox_Department.Items.Add($Name)
                }
				$ComboBox_Department.Enabled = $True
                [System.Windows.Forms.Application]::DoEvents()
              }
           }

    }`
    -UpdateScript {
    }

Add-JobTracker -Name "GetTitles" `
    -JobScript {
        try {
        (get-aduser -Filter "enabled -eq 'true'" -property title).title | Sort-Object -Unique
        }
        catch {}
    }`
    -CompletedScript {
        Param($Job)
           $result = Receive-Job -Job $Job -Keep
           if($result) {
                $array = New-Object System.Collections.ArrayList
                Foreach ($name in $result) {
                    $array.Add($name) > $null
                }
              if($ComboBox_Title) {
				  $ComboBox_Title.Items.clear()
                Foreach ($name in ($array | Sort-Object))
                {
                    [void] $ComboBox_Title.Items.Add($Name)
                }
				$ComboBox_Title.Enabled = $True
                [System.Windows.Forms.Application]::DoEvents()
              }
           }

    }`
    -UpdateScript {
    }

Add-JobTracker -Name "GetAllUsers" `
    -JobScript {
         get-aduser -SearchBase $using:OUPath -Filter "enabled -eq 'true'" # -filter *
    }`
    -CompletedScript {
        Param($Job)
        $result = Receive-Job -Job $Job -Keep
        if($result) {
            $global:AdUsers = $result #Save the Data in an array to be used later on!
        }
    }`
    -UpdateScript {
    }
         

Add-JobTracker -Name "GetAllUPNs" `
    -JobScript {
    $result =@()
        try { # Get the current Active Directory Domain Details
            Import-Module ActiveDirectory
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()            
            $domaindn = ($domain.GetDirectoryEntry()).distinguishedName            
        }
        catch {}
        if($domaindn) {
            $upnDN = "cn=Partitions,cn=Configuration,$domaindn"
            try { # Get all the UPNs in the current domain
                Get-ItemProperty -Path ad:\$upnDN -Name upnsuffixes | select -ExpandProperty upnsuffixes
            }
            catch {}
        }

        
    }`
    -CompletedScript {
        Param($Job)
        $result = Receive-Job -Job $Job -Keep
        if($result) {
            $global:UPNS = $result  #Save the Data in an Array to be used later on!

            $array = New-Object System.Collections.ArrayList
            Foreach ($Name in $result) {
                $array.Add($Name) > $null
            }
            if($ComboBox_DomainName) {
				$ComboBox_DomainName.Items.clear()
                Foreach ($Name in ($array | Sort-Object))
                {
                    [void]$ComboBox_DomainName.Items.Add($Name)
                 }
				$ComboBox_DomainName.Enabled = $True
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }`
    -UpdateScript {
    }

 
[void] $TextBox_User.Focus()
[System.Windows.Forms.Application]::DoEvents()

[void]$main_form.ResumeLayout()

Remove_All_Controls($formLoading)
[void]$formLoading.Close()
[System.Windows.Forms.Application]::DoEvents()
[void]$formLoading.Dispose()
[System.Windows.Forms.Application]::DoEvents()

##[system.windows.forms.application]::run($main_form)
[void]$main_form.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
[System.Windows.Forms.Application]::DoEvents()

if ($main_form.Tag -eq $null) { 
    
    ##$TextBox_User.remove_
    Remove_All_Controls($main_form)
    [void]$main_form.Close()
    [void]$main_form.Dispose()
    Stop-JobTracker
    Get-Event | Remove-Event
    exit

}

$formRunning = New-Object System.Windows.Forms.Form

$labelRunning = New-Object System.Windows.Forms.Label
$labelInfo = New-Object System.Windows.Forms.Label
$okButton = New-Object System.Windows.Forms.Button

  [void] $okButton.Add_Click({ 
        Remove_All_Controls($formRunning)
        [void]$formRunning.Close()
        [void]$formRunning.Dispose()
        Stop-JobTracker
        Get-Event | Remove-Event
        exit
  })

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
  
# Form Events
  $Form_StateCorrection_Load = 
  {

    #Correct the initial state of the form to prevent the .Net maximized form issue
    $formRunning.WindowState = $InitialFormWindowState;
    $hrgn = $Win32Helpers::CreateRoundRectRgn(0,0,$formRunning.Width, $formRunning.Height, 4,4);
    $formRunning.Region = [Region]::FromHrgn($hrgn);

  }
  
  $Form_Cleanup_FormClosed =
  {

    #Remove all event handlers from the controls
    try {
      ##[void] $formRunning.remove_Load($formLoading_Load);
      [void] $formRunning.remove_Load($Form_StateCorrection_Load);
      [void] $formRunning.remove_FormClosed($Form_Cleanup_FormClosed);
    } catch [Exception] { }
  
  }
  
  $formRunning.Font = $Font14B

  $ScriptText = " Creating $($TextBox_User.text)@$($ComboBox_DomainName.Text) - Please Wait... "
  $aSize = [System.Windows.Forms.TextRenderer]::MeasureText($ScriptText, $Font14B)
  [int]$getWidth = $aSize.width + 30
  [int]$getHeight = 92
  [int]$buttonHeight = 32
  
  [void]$formRunning.Controls.Add($labelRunning)
  [void]$formRunning.Controls.Add($labelInfo)
  [void]$formRunning.Controls.Add($okButton)

  $formRunning.BackColor = $BackColor
  $formRunning.ForeColor = $ForeColor
  $formRunning.ControlBox = $False
  $formRunning.Cursor = 'AppStarting'
  $formRunning.FormBorderStyle = 'None'  #'FixedToolWindow'
  $formRunning.Name = "Processing"
  $formRunning.ShowIcon = $False
  $formRunning.ShowInTaskbar = $False
  $formRunning.StartPosition = 'CenterScreen'
  $formRunning.Text = ""
  $formRunning.AutoSize = $true
  $formRunning.ClientSize = "$($getWidth), $($getHeight)"
  
  $labelRunning.Location = '5, 5'
  $labelRunning.Size = "$getWidth, 36"
  $labelRunning.TabIndex = 0
  $labelRunning.AutoSize = $false
  $labelRunning.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $labelRunning.Text = $ScriptText

  $labelInfo.Location = '5, 43'
  $labelInfo.Size = "$getWidth, 30"
  $labelInfo.TabIndex = 0
  $labelInfo.AutoSize = $false
  $labelInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $labelInfo.Text = "Processing..."
  $labelInfo.Font = $Font8
   
  $buttonPosX = $getWidth - 70
  $buttonPosY = ($getHeight - $buttonHeight)

  $okButton.Location = New-Object System.Drawing.Size($buttonPosX, $buttonPosY)
  $okButton.Size = New-Object System.Drawing.Size(60, $buttonHeight)
  $okButton.Text = "OK"
  $okButton.FlatAppearance.BorderColor = $ButtonBorderColor
  $okButton.FlatAppearance.BorderSize = 1
  $okButton.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
  $okButton.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
  $okButton.FlatStyle = "Flat"
  $okButton.ForeColor = $ForeColor
  $okButton.BackColor = $ButtonColor
  $okButton.font = $ButtonFont
  $okButton.visible = $false
  
$InitialFormWindowState = $formRunning.WindowState #Save the initial state of the form
[void]$formRunning.add_Load($Form_StateCorrection_Load) #Init the OnLoad event to correct the initial state of the form  
[void]$formRunning.add_FormClosed($Form_Cleanup_FormClosed) #Clean up the control events
[void]$formRunning.ResumeLayout()
[System.Windows.Forms.Application]::DoEvents()
    

#Show the Form
[void]$formRunning.Show((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
[System.Windows.Forms.Application]::DoEvents()

#Move the form down a litle bit on the screen - unused as there is no authentication window being used
#$formRunning.top = $formRunning.top + 200
#[System.Windows.Forms.Application]::DoEvents()

$errorOccured = $false
$Today = ( get-date -Format("dddd d MMMM yyyy"))

$manager = ""
if($ComboBox_Manager.SelectedItem) { 
    $searchManager = $ComboBox_Manager.SelectedItem.ToLower()
}

if ($searchManager) {
    $User = $global:AdUsers | Where-Object {$_.Name -like "$($searchManager)" }
    if ($User) { 
		if ($CheckBox_UseLaptopOrPC.Checked -eq $true) {
			$manager = $User.DistinguishedName
		} else {
			$manager = $User.UserPrincipalName
		}
	}
}

$department = ""
if($ComboBox_Department.Text) { $department = $ComboBox_Department.Text}
if($ComboBox_Department.SelectedItem) { $department = $ComboBox_Department.SelectedItem}

$password = $TextBox_Password.Text 

if($TextBox_User.Text -EQ "") { 

    [void]$main_form.Dispose()
    Remove_All_Controls($formRunning)
    [void]$formRunning.Close()
    [void]$formRunning.Dispose()
    Stop-JobTracker
    Get-Event | Remove-Event
    exit   #end the script
}

# Create the Email Address first firstname.lastname@domain.org

if($ComboBox_DomainName.Text) { $email = $ComboBox_DomainName.Text }
if($ComboBox_DomainName.SelectedItem) { $email = $ComboBox_DomainName.SelectedItem}

$loginname = $TextBox_User.Text
$firstname = $TextBox_FirstName.Text
$lastname = $TextBox_LastName.Text
$displayname = $TextBox_DisplayName.Text
$password = $TextBox_Password.Text

$SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force
$emailaddress = "$($loginname)@$($email)"
$emaillowercase = $emailaddress.ToLower()
$emailuppercase = $emailaddress.ToUpper()

$phillpines = "N"
$newuser = "N"

if ($CheckBox_UseLaptopOrPC.Checked -eq $true) {
    $newuser = "Y"
}

$NewGroups = ""
$365License = $ComboBox_License.SelectedIndex

#Set Address Details or use AU as default ******************************************************************************************************

$country = "AU"

if($TextBox_Address.Text) {
	$streetaddress = $TextBox_Address.Text
} else {
	$streetaddress = $AUStreetAddress
}

if($TextBox_City.Text) {
	$city = $TextBox_City.Text
} else {
	$city = $AUCity
}

if($TextBox_State.Text) {
	$state = $TextBox_State.Text
} else {
	$state = $AUstate
}

if($TextBox_PostCode.Text) {
	$postalcode = $TextBox_PostCode.Text
} else {
	$postalcode = $AUpostalcode
}

if ($CheckBox_SecondOffice.Checked -eq $true) { 
    $phillpines = "Y"
    $country = "PH"
}

$title = $ComboBox_Title.Text
$phone = $TextBox_Mobile.Text
$directphone = $TextBox_Phone.Text

$created = $false

##################################################### Create user in Office 365 #################################################################################################
if($newuser -EQ "N" ) {

    $labelInfo.Text = "Authenticating via MSGraph to Office 365."
    [System.Windows.Forms.Application]::DoEvents()
        
	#Setup the new Connection in 365 using Microsoft Graph Powershell
    
    #Connect-MgGraph -ContextScope Process -ForceRefresh   #To connect as a different identity other than CurrentUser, specify the -ContextScope parameter with the value Process.
    #Connect-MgGraph -ClientId "YOUR_APP_ID" -TenantId "YOUR_TENANT_ID" -CertificateThumbprint "YOUR_CERT_THUMBPRINT"
    #Connect-MgGraph -AccessToken $AccessToken
    #Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All" -UseDeviceAuthentication     

    #$RequiredScopes = @(Directory.AccessAsUser.All, Directory.ReadWrite.All)
    #Connect-MgGraph -Scopes $RequiredScopes
    #
    #Full Access - Connect-MgGraph -Scopes Directory.AccessAsUser.All, Directory.ReadWrite.All
	
    # To get the Details of the Sigedin User: Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/me
	# Scope         = "https://graph.microsoft.com/.default"
	
    $Authenticated = (Get-MgContext -ErrorAction SilentlyContinue)
    
	$scp = @('Mail.ReadWrite','User.ReadWrite.All','Calendars.Read','Mail.ReadBasic.All','Application.ReadWrite.All','Directory.ReadWrite.All','MailboxSettings.Read','Contacts.ReadWrite','Directory.Read.All','User.Read.All','Organization.ReadWrite.All','Mail.Read','Calendars.ReadWrite','LicenseAssignment.ReadWrite.All','Mail.Send','MailboxSettings.ReadWrite','Organization.Read.All','Contacts.Read','Mail.ReadBasic','Group.ReadWrite.All')
	
    		
	if(-not $Authenticated) {

        try {
            
            $TenantId = $script:Tenant #  # aka Directory ID. This value is Microsoft tenant ID
            $ClientId = $script:AzureClientApp   # aka Application ID
            $ClientSecret = $script:AzureClientPassword  # aka key      VALUE:  7c0c029d-edf6-47bf-ab48-4e9e8f0cc5e2
            #App Registration details
            $Body =  @{
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                Client_Id     = $ClientID
                Client_Secret = $ClientSecret
            }
 
            $Connection = Invoke-RestMethod -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Method Post -Body $body -TimeoutSec 30
			##$Connection = Invoke-RestMethod -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -TimeoutSec 30
		
            #Get the Access Token 
            $Token = $Connection.access_token
			##$Token = ($Connection.Content | ConvertFrom-Json).access_token
						
			# Check if v1.0 or v2.0 of the Microsoft Graph PowerShell module
			$targetParameter = (Get-Command Connect-MgGraph).Parameters['AccessToken']
			if ($targetParameter.ParameterType -eq [securestring]){
				
				Connect-MgGraph -AccessToken ($Token |ConvertTo-SecureString -AsPlainText -Force) > $null
				
			} else {
				
                $connection = Connect-MgGraph -AccessToken $Token > $null
				
			}
			
            Start-Sleep -Seconds 1.5
			
			# $scopes = Get-MgContext | Select-Object -ExpandProperty Scopes 
			# $account = Get-MgContext | Select-Object -ExpandProperty Account
			# if($account) { Write-host $account }
			# if($scopes) { Write-host $scopes }	
			
			
         } catch {

			Write-Host "Authentication with Microsoft Graph Was Canceled or Failed."
			Write-Host $_
			
            [void]$main_form.Dispose()
            [void]$formRunning.Close()
            [void]$formRunning.Dispose()
            Stop-JobTracker
            Get-Event | Remove-Event
            Get-Variable | Remove-Variable -ErrorAction SilentlyContinue
            
			
            exit

         }
		 
		 <#
		 # Use Exchangeonline until you work out how to use msgraph to do the same functions
		 try { 
			Connect-ExchangeOnline -UserPrincipalName (Get-MgContext).Account -ShowBanner:$false
		 } catch {
			Write-Host "Authentication with ExchnageOnline Failed."
			Write-Host $_.Exception.message -ForegroundColor Red	
			
			[void]$main_form.Dispose()
            [void]$formRunning.Close()
            [void]$formRunning.Dispose()
            Stop-JobTracker
            Get-Event | Remove-Event
            Get-Variable | Remove-Variable -ErrorAction SilentlyContinue
			exit
		 }
		 #>
		 
    }

    $labelInfo.Text = "Searching for an existing user mailbox."
    [System.Windows.Forms.Application]::DoEvents()

    try
    {
        $ExistingMailbox = (Get-MgUser -filter "UserPrincipalName eq '$($emaillowercase)'")
        $errorOccured = $False

    } catch {

        $errorOccured = $True

    }
    
	if ($ExistingMailbox) {  # Try Again just incase the old user acoount was just deleted....

	    $labelInfo.Text = "Searching again for an existing Office 365 Mailbox."
		[System.Windows.Forms.Application]::DoEvents()
		Start-Sleep -Seconds 4
		
        try {

			$ExistingMailbox = (Get-MgUser -filter "UserPrincipalName eq '$($emaillowercase)'")
			$errorOccured = $False

		} catch {

			$errorOccured = $True

		}
	}
	
    if (-not $ExistingMailbox -AND $errorOccured -eq $False) {
        
		$labelInfo.Text = "Creating mailbox in location AU."
        [System.Windows.Forms.Application]::DoEvents()

        $NewPasswordProfile = @{}
        $NewPasswordProfile["Password"]= $password
        $NewPasswordProfile["ForceChangePasswordNextSignIn"] = $False
        $hashtable = @{
            UserPrincipalName = $emaillowercase
            DisplayName       = $displayname
            PasswordProfile   = $NewPasswordProfile
            MailNickName      = $loginname
            City              = $city
            CompanyName       = $company
            GivenName         = $firstname
            Surname           = $lastname
            PostalCode        = $postalcode
            State             = $state 
            StreetAddress     = $streetaddress
            Country           = $country
            MobilePhone       = $phone
            BusinessPhones    = $directphone
            JobTitle          = $title
            Department        = $department
            AccountEnabled    = $true
            UsageLocation     = 'AU'
            ErrorAction       = 'Stop'
        }
        
        # Remove Blank Key Values from Hashtable
        $keysToRemove = $hashtable.keys | Where-Object { !$hashtable[$_] }
        $keysToRemove | Foreach-Object { $hashtable.remove($_) }
        
        try {

            $New365user = (New-MgUser @hashtable)
            $errorOccured = $False
			$created = $True
            Start-Sleep -Seconds 1.5

        } catch {

            $message = $_test.user
            $errorMEssage = "ERROR - Could not create a new email account $($emaillowercase) `n$($message)"
            $errorOccured = $True

        }
        
        if($manager) {

			$labelInfo.Text = "Adding a manager to the users account."
			[System.Windows.Forms.Application]::DoEvents()
			Start-Sleep -Seconds 2 # Wait 2 secs to allow cloud to add new user??
			
			$ManagerID = (Get-MgUser -filter "UserPrincipalName eq '$($manager)'").Id
			$UserID = (Get-MgUser -filter "UserPrincipalName eq '$($emaillowercase)'").Id
			
			if($ManagerID -AND $UserID) {
				try {
					$newman=(Set-MgUserManagerByRef -UserId $UserID -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$ManagerID" })
					$errorOccured = $false
				} catch { 
					$errorMEssage = "ERROR - Could add a manager '$($manager)' to the users account."
					$errorOccured = $True
				}

			}
        }

        $labelInfo.Text = "Retrieving Tenant 365 Licences."
        [System.Windows.Forms.Application]::DoEvents()
		
		# This sextion was required for V1.0 MS Graph
        #try {
        #        Select-MgProfile v1.0
        #        $errorOccured = $False
        #    } catch {
        #        $message = $_
        #        $errorMEssage = "ERROR - Could not connect to MgProfile V1.0"
        #        $errorOccured = $True
        #}

        # Visio 2 - SKU - c5928f49-12ba-48f7-ada3-0d743a3601d5
        # Defender 1 - SKU - 4ef96642-f096-40de-a3e9-d83fb2f90211
    
        # E3 - SKU 6fd2c87f-b296-42f0-b197-1e91e994b900  - ENTERPRISEPACK - E3
        # E1 - SKU - 18181a46-0d4e-45cd-891e-60aabd171b4e  - STANDARDPACK - E1
        # E5 - SKU - Unknown Add later if license are purchased
        # 365 - SKU - 3b555118-da6a-4418-894f-7df1e2096870 - O365_BUSINESS_ESSENTIALS -Business Essentials 
        # 365 - SKU - f245ecc8-75af-4f8e-b61f-27d8114de5f3 - O365_BUSINESS_PREMIUM - Business Premium
        # Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq 6fd2c87f-b296-42f0-b197-1e91e994b900)" -All

        #Setup License Calculation Variables

        $Allocate = ""
        [int]$consumed = 0
        [int]$active = 0
        
        try {
            # Get a specific commercial subscription that the organization has acquired
            $MsolSkus = Get-MgSubscribedSKU -All -Property @("SkuId", "SkuPartNumber", "ConsumedUnits", "PrepaidUnits") | Select-Object *, @{Name = "ActiveUnits"; Expression = { ($_ | Select-Object -ExpandProperty PrepaidUnits).Enabled } } | Select-Object SkuId, SkuPartNumber, ActiveUnits, ConsumedUnits
            $errorOccured = $False
        } catch {
            $message = $_
            $errorMEssage = "ERROR - Could not get list of all the current Licenses."
            $errorOccured = $True
        }

        

        $E1License = $MsolSkus| Where-Object {$_.SkuId -EQ "$E1"}
        $E3License = $MsolSkus| Where-Object {$_.SkuId -EQ "$E3"}
        $E5License = $MsolSkus| Where-Object {$_.SkuId -EQ "$E5"}
        $BPLicense = $MsolSkus| Where-Object {$_.SkuId -EQ "$BP"}
        $BELicense = $MsolSkus| Where-Object {$_.SkuId -EQ "$BE"}
        $BDLicense = $MsolSkus| Where-Object {$_.SkuId -EQ "$Defender"} # Business Defender

        $lic_Defender = ""
        $lic_E1 = ""
        $lic_E3 = ""
        $lic_E5 = ""
        $lic_BP = ""
        $lic_BE = ""
        
        [int]$lic_DEF_Consumed = 0
        [int]$lic_E1_Consumed = 0
        [int]$lic_E3_Consumed = 0
        [int]$lic_E5_Consumed = 0
        [int]$lic_BP_Consumed = 0
        [int]$lic_BE_Consumed = 0
        
        [int]$lic_DEF_Available = 0
        [int]$lic_E1_Available = 0
        [int]$lic_E3_Available = 0
        [int]$lic_E5_Available = 0
        [int]$lic_BP_Available = 0
        [int]$lic_BE_Available = 0

        [int]$lic_DEF_Active = 0
        [int]$lic_E1_Active = 0
        [int]$lic_E3_Active = 0
        [int]$lic_E5_Active = 0
        [int]$lic_BP_Active = 0
        [int]$lic_BE_Active = 0

        # Is there enough Defender Licenses?
        if ($BDLicense)
        {

            $lic_DEF_Consumed = [int]$BDLicense.ConsumedUnits
            $lic_DEF_Active = [int]$BDLicense.ActiveUnits
            $lic_DEF_Available = [int]$BDLicense.ActiveUnits - [int]$BDLicense.ConsumedUnits
            
            if ($lic_DEF_Consumed -lt $lic_DEF_Active)
            {
                   $lic_Defender = $Defender
            }
            
        }
        
        # Is there enough E1 Licenses?
        if ($E1License) {

            $lic_E1_Consumed = [int]$E1License.ConsumedUnits
            $lic_E1_Active = [int]$E1License.ActiveUnits
            $lic_E1_Available = [int]$E1License.ActiveUnits - [int]$E1License.ConsumedUnits
            
            if ($lic_E1_Consumed -lt $lic_E1_Active) {
                   $lic_E1 = $E1
            }

        }
        
        # Is there enough E3 Licenses?
        if ($E3License) {

            $lic_E3_Consumed = [int]$E3License.ConsumedUnits
            $lic_E3_Active = [int]$E3License.ActiveUnits
            $lic_E3_Available = [int]$E3License.ActiveUnits - [int]$E3License.ConsumedUnits

            if ($lic_E3_Consumed -lt $lic_E3_Active) {

                   $lic_E3 = $E3

            }
        }       
        
        # Is there enough E5 Licenses?
        if ($E5License) {

            $lic_E5_Consumed = [int]$E5License.ConsumedUnits
            $lic_E5_Active = [int]$E5License.ActiveUnits
            $lic_E5_Available = [int]$E5License.ActiveUnits - [int]$E5License.ConsumedUnits

            if ($lic_E5_Consumed -lt $lic_E5_Active) {
                   $lic_E5 = $E5
            }
        }       
        
        # Is there enough Business Essential Licenses?
        if ($BELicense) {

            $lic_BE_Consumed = [int]$BELicense.ConsumedUnits
            $lic_BE_Active = [int]$BELicense.ActiveUnits
            $lic_BE_Available = [int]$BELicense.ActiveUnits - [int]$BELicense.ConsumedUnits

            if ($lic_BE_Consumed -lt $lic_BE_Active) {
                   $lic_BE = $BE
            }
        }       
        
        # Is there enough Business Premium Licenses?
        if ($BPLicense) {

            $lic_BP_Consumed = [int]$BPLicense.ConsumedUnits
            $lic_BP_Active = [int]$BPLicense.ActiveUnits
            $lic_BP_Available = [int]$BPLicense.ActiveUnits - [int]$BPLicense.ConsumedUnits

            if ($lic_BP_Consumed -lt $lic_BP_Active) {
                   $lic_BP = $BP
            }
        }       

        # Workout what License to allocate to the new mailbox
        if( $365License -eq 0 ) {

            $LIC = "E1"
            $consumed = $lic_E1_Consumed
            $active = $lic_E1_Active

            if($lic_E1) {

                $Allocate = $lic_E1
           
            }
        }
        
        if( $365License -eq 4 ) {

            $LIC = "Business Premium"
            $consumed = $lic_BP_Consumed
            $active = $lic_BP_Active

            if($lic_BP) { $Allocate = $lic_BP }
        }

        if( $365License -eq 1 ) {

            $LIC = "E3"
            $consumed = $lic_E3_Consumed
            $active = $lic_E3_Active

            if($lic_E3) { $Allocate = $lic_E3 }
        }

        if( $365License -eq 2 ) {
            $LIC = "E5"
            $consumed = $lic_E5_Consumed
            $active = $lic_E5_Active
            if($lic_E5) { $Allocate = $lic_E5 }
        }
        if ($lic_Defender) {
            $DEF = "+ Defender"
        } else {
            $DEF = ""
        }

        # Have worked out what License to use - If so then Allocated it and also a Defender License if there is engough left?
        if ($Allocate) {

            if ($lic_Defender) {

               $LicenseParams = @{
                   AddLicenses = @(
                       @{
                           DisabledPlans = @()
                           SkuId = "$($lic_Defender)" # Defender 1
                       }
                       @{
                           DisabledPlans = @()
                           SkuId = "$($Allocate)" # E1 / BP
                       }
                   )
                    RemoveLicenses = @()
               }
            } else   {
              $extraMessage =  "ERROR - Not enough Defender (Plan 1) licenses available. `nPurchase New Licenses via RHIPE ( https://www.prismportal.online ) `nUsed: $($lic_DEF_Consumed)"

              $LicenseParams = @{
                  AddLicenses = @(
                      @{
                          DisabledPlans = @()
                          SkuId = "$($Allocate)" # E1 / BP
                      }
                  )
                   RemoveLicenses = @()
              }
            } # Defender License

        } else {

            # NO Allocate
            $extraMessage =  "ERROR - Not enough $LIC licenses available. `nPurchase New Licenses via RHIPE ( https://www.prismportal.online )`nActive: $($active)  Used: $($consumed)`nManually Add a new License to this new user using the Office 365 ADMIN Portal."
            
        } # Allocate

    
		# Allocate the Licenses to the Email Account         
		$NewMailbox = $null
		
		if($errorOccured -eq $False) {

			$loops = 0
			do {
				$labelInfo.Text = "Waiting for Azure Account to be Created... ($loops)"
				[System.Windows.Forms.Application]::DoEvents()
				
				try {
					$NewMailbox =  Get-MgUser -filter "UserPrincipalName eq '$($emaillowercase)'"
				} catch { }
				if($newMailbox) {
					$loops = 10
					$errorOccured = $False
				}
				if($loops -le 4) {
					Start-Sleep -Seconds 10 # Wait 10 secs to allow cloud to add new user??
				}
				$loops++
			} while ($loops -le 3)
        

			if ($NewMailbox) {
				if ($lic_Defender) {
					$labelInfo.Text = "Allocating Defender and $LIC Licenses."
					[System.Windows.Forms.Application]::DoEvents()
				} else {
					$labelInfo.Text = "Allocating a $LIC License."
					[System.Windows.Forms.Application]::DoEvents()
				}
	
				try {
					$NewLic = (Set-MgUserLicense -UserId "$($emaillowercase)" -BodyParameter $LicenseParams)
					$errorOccured = $False
				} catch {
					$message = $_
					$errorMessage =  "ERROR - Allocating a New $LIC License - FAILED `nTry Adding a License to the email account manually via the Office ADMIN Portal. `n$($message)"
					$errorOccured = $True
				}
				
				Start-Sleep -Seconds 1
				
				$UserID = $NewMailbox.id
				$UserPN = $NewMailbox.UserPrincipalName
				
				<#
				#Update the Users Timezone once the Mailbox has been created.
	
				#Get a list of timezones form the Registry
				#Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | Format-List pschildname
				#$User = Get-MgUser -filter "UserPrincipalName eq '$($emaillowercase)'" -ConsistencyLevel eventual
				
				$labelInfo.Text = "Updating Mailbox Timezone. (For $UserPN)"
				[System.Windows.Forms.Application]::DoEvents()
				

				if($CheckBox_SecondOffice.Checked -eq $True ) {
				
						$params = @{
							"@odata.context" = "https://graph.microsoft.com/v1.0/$metadata#Me/mailboxSettings"
							dateFormat ="dd-MM-yyyy"
							timeZone = "Singapore Standard Time"
							timeFormat = "h:mm tt"
						}
					try {
					
						Update-MgUserMailboxSetting -UserId $UserPN -BodyParameter $params | Out-Null
					
					} catch {
						$message = $_
						$errorMessage =  "ERROR - Updating Mailbox Timezone. `n$($message)"
						$errorOccured = $True
					}
		
				} else {
						$params = @{
							"@odata.context" = "https://graph.microsoft.com/v1.0/$metadata#Me/mailboxSettings"
							dateFormat ="dd/MM/yyyy"
							timeZone = "E. Australia Standard Time"
							timeFormat = "h:mm tt"
						}
						
					try {
					
						#$NewTZ = (Update-MgUserMailboxSetting -UserId "$($emaillowercase)" -DateFormat "dd/MM/yyyy" -TimeZone "E. Australia Standard Time" -TimeFormat "h:mm tt")
						Update-MgUserMailboxSetting -UserId $UserPN -BodyParameter $params | Out-Null
						
						
					} catch {
						$message = $_
						$errorMessage =  "ERROR - Updating Mailbox Timezone. `n$($message)"
						$errorOccured = $True
					}
				}
				
				Start-Sleep -Seconds 1
				#>
				
				# Add a User to an Azure AD Group with the GroupID and BodyParameter Parameters
				# $UserUPN = (Get-MgUser | Where-Object {($_.DisplayName -eq 'Victor Ashiedu') -and ($_.UserPrincipalName -like '*@itechguides.com*')}).UserPrincipalName
				# $BodyParams = @{
				# "@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserUPN"
				# }
				# $Groupid = (Get-MgGroup | Where-Object {$_.DisplayName -eq "NewSecurityGroup"}).id
				# New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $BodyParams
				# Get-MgGroupMember -GroupId $GroupId
				
				
				$UserBased = $null
				$allgroups = $null
				
				if($ComboBox_BasedOn.SelectedItem) { 
					$UserBased= $ComboBox_BasedOn.SelectedItem.ToLower()
				}
				if($UserBased) {
					$labelInfo.Text = "Adding User Like $($UserBased) Group(s)."
					[System.Windows.Forms.Application]::DoEvents()
					$UserBasedOn = (Get-MgUser -filter "DisplayName eq '$UserBased'" -ConsistencyLevel eventual)
					Start-Sleep -Seconds 1
					if($UserBasedOn) {
						$userbasedonpn = $UserBasedOn.UserPrincipalName
						$basedonUPN = $UserBasedOn.UserPrincipalName
						
						$BasedOn = (Get-MgUserMemberOf -UserId $basedonUPN -All -ConsistencyLevel eventual | Select-Object *)
						#$existing = (Get-MgUserMemberOf -UserId $UserPN -All -ConsistencyLevel eventual | Select-Object *).id
						
						if($BasedOn) {
							$labelInfo.Text = "Retrieving All Azure Groups."
							[System.Windows.Forms.Application]::DoEvents()
							
							try {
								$allGroups = Get-MgGroup -Filter "groupTypes/any(x:x eq 'unified')" -All
							} catch {}
							Start-Sleep -Seconds 1

							if($allGroups) {
								foreach ($groupId in $BasedOn) {
									$gid = $groupId.id
									$group = $allGroups | Where-Object { $_.id -eq $gid }
									
									if($group) {
										$groupName = $group.DisplayName
										$groupID = $group.id
										
										$labelInfo.Text = "Adding user to Group $groupName"
										[System.Windows.Forms.Application]::DoEvents()
																					
										try {
											$newmember=(New-MgGroupMember -GroupId $groupID -DirectoryObjectId $UserID | Out-Null)
											$errorOccured = $False
										} catch {
											$message = $_
											$errorMessage =  "ERROR - Adding User to Azure Group $groupName. `n$($message)"
											$errorOccured = $True
										}	
										if($errorOccured -eq $False) {
											if($NewGroups) {
												$NewGroups = $NewGroups + ", " +$groupName
											} else {
												$NewGroups = $groupName
											}
										}										
										Start-Sleep -Seconds 1
									}
								}
							}
						
							$connectedSP = $False
							$labelInfo.Text = "Office 365 Sharepoint Site $SPSite"
							[System.Windows.Forms.Application]::DoEvents()
							
							$AuserName = $script:ConnectSPOServiceUser
							$aPassword = $script:ConnectSPOServicePassword
							
							$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AuserName, $(convertto-securestring $aPassword -asplaintext -force)
								
							try {
								Connect-SPOService -url "$SPAdminSite" -Credential $cred | Out-Null
								$connectedSP = $True
							} catch {}
						
							if($connectedSP -eq $True) {
								$SPUSer = $null
								$loops = 1
								do {
									$labelInfo.Text = "Waiting for 365 Sharepoint user $UserPN. [$loops]"
									[System.Windows.Forms.Application]::DoEvents()
									
									try {
										$SPUSer = (Get-SPOUser -LoginName $UserPN -Site $SPSite)
									} catch { }
									if($SPUSer) {
										$loops = 10
										$errorOccured = $False
									}
									if($loops -le 4) {
										Start-Sleep -Seconds 10 # Wait 10 secs to allow cloud to add new user??
									}
									$loops++
								} while ($loops -le 4)
								
								$Users = @()
								
								if($SPUSer) {
									
									$CurrentGroups = $SPUser.Groups
									
									foreach($SPgroup in $SharepointGroups) {
										
										$labelInfo.Text = "Checking 365 Sharepoint Group '$SPgroup'"
										[System.Windows.Forms.Application]::DoEvents()
										
										$Users = (Get-SPOSiteGroup -Site $SPSite -Group $SPgroup).Users
										
										if($Users -Contains ($userbasedonpn)) {

											if($SPUSer.Groups -NOTContains($SPgroup)) { 
												$labelInfo.Text = "Adding user to Group '$SPgroup'"
												[System.Windows.Forms.Application]::DoEvents()
												
												try {
													Add-SPOUser -Group $SPgroup -Site $SPSite -LoginName $UserPN -ErrorAction SilentlyContinue | Out-Null
													$errorOccured = $False
												} catch {
													$message = $_
													$errorMessage =  "ERROR - Adding $UserPN to SharePoint Group '$SPgroup' `n$($message)"
													$errorOccured = $True
												}
												if($errorOccured -eq $False) {
													Start-Sleep -Seconds 1
													if($NewGroups) {
														$NewGroups = $NewGroups +", " +$SPgroup
													} else {
														$NewGroups = $SPgroup
													}

												}
											}# IF User already Member of group (somehow?)
										} # IF CheckUser is a member of the Group
									} #For Each
								} #IF SPUser
								Disconnect-SPOService| Out-Null
							}
						}
					}
				}
				
				Start-Sleep -Seconds 1
				
			} else {

				$errorMessage =  "ERROR - Cant allocate license as the new account was never created?"
				$errorOccured = $True
			}

		} # errorOcured 

	} else { # User Mailbox already Exists in 365.

         $errorMessage = "FAILED - The account $($emaillowercase) already exists in Office 365!`nPlease re-create them again using a unique Email Address."
         $errorOccured = $True

	}

}


##################################################### Create user in Active Directory #########################################################################################
if($newuser -EQ "Y") {

    $labelInfo.Text = "Creating $($loginname) in Active Directory."
    [System.Windows.Forms.Application]::DoEvents()

    #Setup the new Connection in Internal AD first - it will get replcated to 365 every 30 mins

    $CheckAD =""
    $CheckAD = $global:AdUsers | Where-Object {$_.SamAccountName -like "$($loginname)" }

    if($CheckAD) {

        $errorMessage = "ERROR - User Already Exists! `nSAM: $($CheckAD.SamAccountName) `nDN: $($CheckAD.DistinguishedName) `nSID: $($CheckAD.SID) "
        $errorOccured = $true

    } else {

        $DName = "$($displayname)"
        $FName = "$($firstname)"
        $LName = "$($lastname)"

        $ADusername = $script:ActiveDirectoryUsername
        $ADpassword = $script:ActiveDirectoryPassword
        $ADsecurePassword = ConvertTo-SecureString $ADpassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($ADusername, $ADsecurePassword)
		
		$proxyAddresses = "SMTP:$($emaillowercase)"
		
        $hashtable = @{
			SamAccountName    = $loginname
            UserPrincipalName = $emaillowercase
            DisplayName       = $displayname
		    Name       		  = $displayname
            GivenName         = $firstname
            Surname           = $lastname
            AccountPassword   = $SecurePassword
            StreetAddress     = $streetaddress
            City              = $city
            Company           = $company
            PostalCode        = $postalcode
            State             = $state 
            Country           = $country
            MobilePhone       = $phone
            OfficePhone       = $directphone
            Title             = $title
            HomePage          = $WWW
            EmailAddress      = $emaillowercase
            Department        = $department
            Description       = "Created $($Today)"
            Enabled           = $true
            Path              = $OUPath
            Manager           = $manager
			ChangePasswordAtLogon = $false
            PasswordNeverExpires  = $true
			CannotChangePassword  = $false
			OtherAttributes       = @{'proxyAddresses' = $proxyAddresses}
        }
        
        # Remove Blank Key Values from Hashtable. IE: If an option is not set then it wont be passed.
        $keysToRemove = $hashtable.keys | Where-Object { !$hashtable[$_] }
        $keysToRemove | Foreach-Object { $hashtable.remove($_) }

        $scriptBlock = {
            param(
                [Parameter(Mandatory=$True, Position=1)]
                [hashtable]$aHashTable
                )
            try {
                    $NewADuser = (New-ADUser @aHashTable)
                } catch {
                    return $_
            }
            
			return $true
        }
			
        $message = ""
		$result = ""
		  
        try {

		    $result = (Invoke-Command -ComputerName $domserver -Credential $credential -ScriptBlock $ScriptBlock -ArgumentList ([hashtable]$hashtable))
		    $errorOccured = $false
			$created = $True
            } catch {

                $message = $_
                $errorOccured = $true
        }
           
        if ($errorOccured -eq $true) {
            $errorMessage = "ERROR - New user failed to be created in active directory. `n$($message)"
        }
		
		if ($result -ne $true) {
            $errorOccured = $true   
            $errorMessage = "ERROR - New user failed to be created in active directory. `n$($result)"

        }
		
		$allstaffgroup = $false	
		$outlookgroup = $false	
		
		if ($result -eq $true) {
			
			$UserBased = $null
				
			if($ComboBox_BasedOn.SelectedItem) { 
				$UserBased = $ComboBox_BasedOn.SelectedItem.ToLower()
			}
		
			if($UserBased) {
			
				$labelInfo.Text = "Copying $($UserBased) Groups."
				[System.Windows.Forms.Application]::DoEvents()
				$UserBasedOn = (Get-ADUser -filter "DisplayName -eq '$UserBased'")
			
				if($UserBasedOn) {
	
					Start-Sleep -Seconds 1
					$usersname = $UserBasedOn.SamAccountName
				
					#Only interested in the checkusers Global Groups
					$BasedOn = (Get-ADPrincipalGroupMembership -Identity $usersname -Server $domserver | where-object {$_.GroupScope -eq 'Global'})
					
					if($BasedOn) {

						foreach ($group in $BasedOn) {
							$groupName = $group.name
							$labelInfo.Text = "Adding User to AD Group $groupName"
							[System.Windows.Forms.Application]::DoEvents()
																			
							$addthisgroup = AddUserToGroup "$groupName" $loginname $domserver
						
							if($addthisgroup -eq $True) {
								Start-Sleep -Seconds 1
								if($NewGroups) {
									$NewGroups = $NewGroups + ", " +$groupName
								} else {
									$NewGroups = $groupName
								}
							}	
						}
					
						# Dont care to capture any group membership error messages
						$errorOccured = $False
					}
				}
			} else {
		
				# Add user to the All Staff and Signatures Group as a bare minumum...?
				if($AllStaffSigGroup -ne "") { $allstaffgroup = AddUserToGroup $AllStaffSigGroup $loginname $domserver }
				if($OutlookSigGroup -ne "") { $outlookgroup = AddUserToGroup $OutlookSigGroup $loginname $domserver }
				
				if ($allstaffgroup -eq $true) {
					if($NewGroups) {
						$NewGroups = $NewGroups + ", $($AllStaffSigGroup)"
					} else {
						$NewGroups = $AllStaffSigGroup
					}
				}
				if ($outlookgroup -eq $True) {
					if($NewGroups) {
						$NewGroups = $NewGroups + ", $($OutlookSigGroup)"
					} else {
						$NewGroups = $OutlookSigGroup
					}
				}
		
			}
		}
	
        $extraMessage = "Manually allocate the appropriate licenses via the 365 ADMIN Portal."

    }
         
    # $errorOccured = $true
}

if($CheckBox_SendEmail.Checked -eq $false) { $sendmail = "" }

$sentmail = $false

if($sendmail -ne "" -and $created -eq $True) {

    $tableuser = New-Object system.Data.DataTable "TableUser"
    $col1 = New-Object system.Data.DataColumn "Info",([string])
    $col2 = New-Object system.Data.DataColumn "Details",([string])
    
    # Add the Columns to the table
    $tableuser.columns.add($col1)
    $tableuser.columns.add($col2)
    
	# Add each Row to the table
    $row = $tableuser.NewRow()
    $row."Info" = "Email"
    $row."Details" = "$($emaillowercase)"
    $tableuser.Rows.Add($row)
    
    $row = $tableuser.NewRow()
    $row."Info" = "Password"
    $row."Details" = "$($TextBox_Password.text)"
    $tableuser.Rows.Add($row)

    $row = $tableuser.NewRow()
    $row."Info" = "FirstName"
    $row."Details" = "$($firstname)"
    $tableuser.Rows.Add($row)

    $row = $tableuser.NewRow()
    $row."Info" = "LastName"
    $row."Details" = "$($lastname)"
    $tableuser.Rows.Add($row)
    
	if($title) {
		$row = $tableuser.NewRow()
		$row."Info" = "Title"
		$row."Details" = "$($title)"
		$tableuser.Rows.Add($row)
	}
	
	if($department) {
		$row = $tableuser.NewRow()
		$row."Info" = "Department"
		$row."Details" = "$($department)"
		$tableuser.Rows.Add($row)
	}
	
	if($searchManager) {
		$row = $tableuser.NewRow()
		$row."Info" = "Manager"
		$row."Details" = "$($searchManager)"
		$tableuser.Rows.Add($row)
	}
	
    $row = $tableuser.NewRow()
    $row."Info" = "Address"
    $row."Details" = "$streetaddress $city $state $postalcode"
    $tableuser.Rows.Add($row)

    $row = $tableuser.NewRow()
    $row."Info" = "Country"
    $row."Details" = "$country "
    $tableuser.Rows.Add($row)

	if($phone) {
		$row = $tableuser.NewRow()
		$row."Info" = "Mobile"
		$row."Details" = "$($phone)"
		$tableuser.Rows.Add($row)
	}
	
	if($directphone) {
		$row = $tableuser.NewRow()
		$row."Info" = "Phone"
		$row."Details" = "$($directphone)"
		$tableuser.Rows.Add($row)
	}
	
    if($newuser -eq "N") {

        $row = $tableuser.NewRow()
        $row."Info" = "License"
        $row."Details" = "$LIC $DEF"
        $tableuser.Rows.Add($row)

    } else {

        $row = $tableuser.NewRow()
        $row."Info" = "License"
        $row."Details" = "** Please add the Licenses manually after the next AzureSync."
        $tableuser.Rows.Add($row)
		
		$row = $tableuser.NewRow()
        $row."Info" = "Sharepoint"
        $row."Details" = "** Please add this user to the Sharepoint Permissions Groups after the next AzureSync."
        $tableuser.Rows.Add($row)

    }

	if ($NewGroups) {
		$row = $tableuser.NewRow()
        $row."Info" = "Groups"
        $row."Details" = "$($NewGroups)"
        $tableuser.Rows.Add($row)
	}
	
    if($newuser -eq "Y") {      

        $aSubject ="New $($company) Active Directory User - $($emaillowercase)"

    } else {

        $aSubject ="New $($company) Microsoft Office 365 User - $($emaillowercase)"

    }

    [string]$body = [PSCustomObject]$tableuser | select -Property "Info", "Details" | ConvertTo-HTML -head $head -PreContent "<font color=`"Black`"><h4>$($aSubject)</h4></font>"
    $body += "<br>"
	
    if($message) {
		$body += "$($message)<br>"
	}
	$body += "<small>"
    if($newuser -EQ "N") {

        if($BPLicense ) { $body += " There are $lic_BP_Available Business Premium licenses available from a total of $lic_BP_Active purchased. <br>" }
        if($BELicense ) { $body += " There are $lic_BE_Available Business Essentials licenses available from a total of $lic_BE_Active purchased. <br>" }
        if($E1License ) { $body += " There are $lic_E1_Available E1 licenses available from a total of $lic_E1_Active purchased. <br>" }
        if($E3License ) { $body += " There are $lic_E3_Available E3 licenses available from a total of $lic_E3_Active purchased. <br>" }
		if($E5License ) { $body += " There are $lic_E5_Available E5 licenses available from a total of $lic_E5_Active purchased. <br>" }
		if($BDLicense ) { $body += " There are $lic_DEF_Available Defender licenses available from a total of $lic_DEF_Active purchased. <br>" }
    
    }
    
    $body += "</small>"
    
    try {
          Send-MailMessage -To $sendmail -From "newuser@$($EmailDomain)" -Subject $aSubject -BodyAsHtml $Body -SmtpServer $SMTPServer -Port 25 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue > $null
          $sentmail = $true
        } catch { $sentmail = $false }

}

# Leave a message with the results
if ($errorOccured -eq $false) { 

    $message = "Sucessfully created - OK"
    if($sentmail -eq $true -and $sendmail) {$message = "$($message) `nConfirmation email sent to $lcsendmail - OK" }
    if($sentmail -eq $false -and $sendmail) {$message = "$($message) `nConfirmation email sent to $lcsendmail - FAILED?" }
    if($extraMessage) { $message = "$($message)`n$($extraMessage)" }

    $labelInfo.Text = $message
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Seconds 5

} else { 

    $labelInfo.Text = $errorMessage
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Seconds 10

}

try {
	Disconnect-MgGraph -ErrorAction Ignore > $null
} catch {}

Stop-JobTracker
[void]$main_form.Dispose()
Remove_All_Controls($formRunning)
[void]$formRunning.Close()
[void]$formRunning.Dispose()

try {
	Remove-Module -Name ActiveDirectory -Force > $null
	Remove-Module -Name Microsoft.Graph.Users -Force > $null
	Remove-Module -Name Microsoft.Graph.Groups -Force > $null
	Remove-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force > $null
	Remove-Module -Name ExchangeOnlineManagement -Force > $null
	Remove-Module -Name Microsoft.Online.SharePoint.PowerShell -Force > $null
	Remove-Module -Name AzureAD -Force > $null
	} catch {}
	

#$runspace.powershell.EndInvoke($runspace.Runspace) > $null
#$runspace.powershell.Runspace.Dispose() # remember to clean up the runspace!
#$runspace.powershell.dispose()

<#

 Start Word Object
$Word = New-Object -ComObject Word.Application

# Optional to make Word visible
$Word.Visible = $True

# Open Word doc
$OpenFile = $Word.Documents.Open("C:\input\Document1.docx")

# Get the content of the doc
$Content = $OpenFile.Content
# My name is $name and the reason is $reason.

# New variable for new text and variables to to replace the ones from the doc.
$newText = ""
$name = "John Doe"
$reason = "testing"

# Store the current text in the var
$newText = $Content.Text

# Replace the template vars with the new values
$newText = $newText  -replace '\$name', $name
$newText = $newText  -replace '\$reason', $reason

# Make the modified text the new content and Save
$Content.Text = $newText
$OpenFile.Save()

#>


