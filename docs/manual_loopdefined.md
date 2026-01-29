# loop-defined 모드 사용법

# loop-defiend 모드
고정된 디렉토리 구조가 아닌 다양한 구조에 대응하여 Plot들을 찾아 슬라이드에 자동으로 삽입하는 모드입니다.

# 디렉토리
## Input Directory
plot을 삽입할 디렉토리를 선택합니다.
만약 디렉토리가 구조가있다면 가장 상위디렉토리를 선택합니다
```
예시 1
📁 pi_case1 <- **이 디렉토리 선택**
├── 📁 C
│   └── 📕 eDep_C_energy{20,40,60,80,100,120}.pdf
├── 📁 DRcor_LCATTcor
│   └── 📕 eDep_DRcor_LCATTcor_energy{20,40,60,80,100,120}.pdf
├── 📁 S_LCATTcor
│   └── 📕 eDep_S_LCATTcor_energy{20,40,60,80,100,120}.pdf
├── 📕 c_piLinearity_M5-T2.pdf
├── 📕 c_piResol_M5-T2.pdf
├── 📕 c_piResol_M5-T2_NNR_From30.pdf
└── 📕 c_piResol_M5-T2_NNR_FullFit.pdf
```

```
예시 2
└── 📁 MuMu
    ├── 📁 UL2016PostVFP 
    │   └── 📁 StackedKinematics <- **이 디렉토리 선택**
    │       ├── 📁 Num_Jets
    │       ├── 📁 Observable
    │       ├── 📁 afterTopReco
    │       │   ├── 📁 Energy
    │       │   ├── 📁 Mass
    │       │   ├── 📁 Rapidity
    │       │   ├── 📁 eta
    │       │   ├── 📁 phi
    │       │   └── 📁 pt
    │       ├── 📁 initial
    │       │   ├── 📁 PV
    │       ├── 📁 step1
    │       │   ├── 📁 Mass
    │       │   ├── 📁 Num_Jets
    │       │   ├── 📁 PV
    │       │   ├── 📁 eta
    │       │   │   ├── 📕 h_{Lep1,Lep2,Jet1,Jet2,bJet1,bJet2}_eta_1.pdf
    │       │   ├── 📁 phi
    │       │   │   ├── 📕 h_{Lep1,Lep2,Jet1,Jet2,bJet1,bJet2}_phi_1.pdf
    │       │   └── 📁 pt
    │       │       ├── 📕 h_{Lep1,Lep2,Jet1,Jet2,bJet1,bJet2}_pt_1.pdf
    │       ├── 📁 step2 (step1과 같은 구조)
    │       ├── 📁 step3 (step1과 같은 구조)
    │       ├── 📁 step4 (step1과 같은 구조)
    │       ├── 📁 step5 (step1과 같은 구조)
    │       └── 📁 step6 (step1과 같은 구조)
```
# loop-defined Options
## Template 사용
- 만약 plot이 저장된 디렉토리가 일정한 구조를 따른다면(예시2) 경로 템플릿을 사용하여 plot을 삽입할수있습니다.
- 파일 이름도 일정 규칙을 따라 생성되었다면 아래와 같이 템플릿화 할 수 있습니다.
```
Path Template: {step}/{kinematic}
Filename Template: h_{object}_{kinematic}_{step}.{format}
```
step: {step1, step2, step3, step4, step5, step6}
kinematic: {pt, eta, phi}
object: Lep1,Lep2,Jet1,Jet2,bJet1,bJet2
로 구성할수있습니다.
## Scan mode 사용
- 혹은 파일명에 따라 Input directory에 있는 모든 파일들을 검색하여 찾을 수도 있습니다.
- 파일명에따라 Variables에 Filename Template Variables에 입력필드가 생기며 이곳에 각각의 변수들을 설정합니다.
(비활성화된 경우 Path Template Variables와 같기 때문이니 변수명을 변경해주세요)
```
h_{ob}_{kinematic}_{num}.{format}
```
이 경우에
ob: Lep1,Lep2,Jet1,Jet2,bJet1,bJet2
kinematic: pt, eta, phi
num: 1,2,3,4,5,6
을 입력한 뒤 **SCAN** 버튼을 누르면 검색결과가 팝업창으로 나옵니다.
팝업창에서 제대로 찾았다면 Correct를 누르고 에러가 났다면 파일명과 변수명을 확인해주세요.
Correct 버튼을 눌러야지 Preview가 활성화 됩니다.

# Preview
- 예를들어 Lep1, Lep2, Jet1, Jet2의 pt, eta, phi를 2*3 배열로 삽입하려면 다음과 같이 합니다.
```
row1: Lep1, Jet1
row2: Lep2, Jet2
column: pt, eta, phi
Slide axis: num
```
미리보기 박스에 어떤 플랏이 삽입될지 보여집니다.
첫번째 슬라이드에는
Lep1 pt | Lep1 eta | Lep1 phi
Lep2 pt | Lep2 eta | Lep2 phi
이렇게 삽입되고
두번째 슬라이드에는
Jet1 pt | Jet1 eta | Jet1 phi
Jet2 pt | Jet2 eta | Jet2 phi
이 순서로 삽입되며 num의 갯수에 따라 6*2장이 생성됩니다.
프리뷰 탭에서 각 페이지별로 plot의 순서를 드래그 앤 드랍으로 조정할 수 있습니다.

## DRC plot
DRC plot를 체크하면 DRC energy distribution을 순서대로 자동삽입합니다.
같은 방식으로 설정합니다
- 디렉토리 구조 있을시
(예시1)
Input Directory: pi_case1
Path Template: {ch}
Filename Template: eDep_{ch}_energy{e}_run{e}.{format}
그런 다음 Variable을 설정합니다.
ch: c, s_lcattcor, drcor_lcattcor
Variable에서는 대소문자를 가리지 않습니다.

- Scan 모드
'Use SCAN mode' 체크 후 Filename Template와 Variable의 Filename Template Variables만 설정합니다.(마찬가지로 대소문자를 가리지 않습니다.)

### DRC plot: review
TB2025에서 받은 Energy point(20,40,60,80,100,120GeV)를 기본으로 2*3배열로 진행합니다.
Slide axis: ch (각 채널 C, S_LCATTcor, DRcor_LCATTcor별로 energy distribution 슬라이드 생성)
Grid axis: e
프리뷰 박스는 드래그 앤 드랍으로 순서를 조정할 수 있습니다.
(만약 e: 20,40,,80,100,120으로 설정되면 60은 건너뜁니다.)
같은 방식으로 프리뷰 탭에서 각 슬라이드별로 plot의 순서를 드래그 앤 드랍으로 조정할 수 있습니다.

# Preset
사용자는 plot 삽입에 필요한 variable들을 preset으로 저장하고 이를 사용하여 빠르게 슬라이드를 만들 수 있습니다.
loop-defined 모드일때 오른쪽에 폴더 모양 아이콘을 누르면 preset을 저장할 위치를 선택합니다.
**폴더를 지정하지 않으면 macOS 표준 경로인 '/Users/{username}/Library/Application Support/SildeMaker'에 저장됩니다.**
처음 사용시, plot을 삽입할때 필요한 variable들을 입력하고 삽입 위치등을 조정합니다.
(반드시 한번 실행해서 제대로 삽입되는지 확인하세요)
그 뒤 드랍다운의 'Save as..'를 눌러 preset의 이름을 적어 저장합니다.
이후에는 드랍다운에서 원하는 설정을 누르면 자동으로 적용되며 바로 plot을 삽입할수 있습니다.
(Scan 모드인 경우 SCAN을 먼저 하세요)

# STOP
plot이 잘못 삽입되고 있는 경우 **STOP**을 누르면 정지됩니다.