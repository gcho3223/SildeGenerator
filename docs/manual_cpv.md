# CPV 모드 사용법

## CPV 모드
ttbar CPV Analysis에서 생성된 Plot들을 슬라이드에 자동으로 삽입하는 모드입니다. 
삽입할 Object와 Event selection Step을 선택합니다.
### **주의: 이 모드는 디렉토리 구조 및 파일 이름이 나한테 맞춰져있음니다**

## 1. Option
### 1) Year
- UL2016PreVFP
- UL2016PostVFP
- UL2017
- UL2018
- Run3 (2022, 2023, 2024)는 추후 업데이트 예정

### 2) Channel
- Dimuon (MuMu)
- Dielectron (ee)
- EMu (emu)
- tau channel은 추후 업데이트 예정

### 3) File Format
- PDF, PNG 파일을 지원합니다.
- JPG, JPEG 등 다른 사진파일들은 필요시 업데이트 예정

## 2. Directories
- Input Directory: 삽입할 Plot이 있는 디렉토리를 선택합니다.
- Output File: Plot들을 삽입한 키노트파일을 저장할 공간을 선택합니다.

## 3. Object
삽입하고자 하는 object를 선택합니다:

### Event selection stage
- **Num_PV**: # of primary vertex
- **Mass**: dilepton invariant mass
- **Num_Jets**: # of jet, # of b jet
- **Lep1, Lep2**: Leading/Sub-leading lepton
    - pT, eta, phi distribution
- **jet1, jet2**: Leading/Sub-leading jet
    - pT, eta, phi distribution
- **bjet1, bjet2**: Leading/Sub-leading b-tagged jet
    - pT, eta, phi distribution
- **MET** : pT, phi distribution
- **Observable**: O1, O3

### After Top quark reconstruction stage 
top quark를 reconstruction 한 뒤의 object들
- **Nu, AnNu**: Energy distribution
- **bJet, AnbJet**: pT, eta, phi energy distribution
- **Top, AnTop**: pT, rapidity, phi, mass distribution

## 4. Event Selection
- **Inital stage**: Noise filter, Trigger까지 통과했을때의 primary vertex distribution을 삽입합니다.
- **step 1**: Dilepton mass cut
- **step 2**: Z mass(peak) veto
- **step 3**: # of Jets ≥ 2
- **step 4**: MET cut
- **step 5**: b-tagging jet ≥ 1
- **setp 6**: top reconstruction

## 5. Systematic 모드 -> 현재 업데이트중
Systematic 체크박스를 활성화하면:
- 여러 샘플(TTbar_Signal, DY, TTV 등)의 비교 플랏을 생성합니다
- Central vs Up vs Down 비교가 가능합니다 
- 각 kinematic 변수당 하나의 슬라이드에 모든 샘플을 배치합니다

## 파일 경로와 파일명 규칙
CPV 모드에서는 다음 경로 규칙을 따릅니다:
```
{input_dir}/{jobversion}/{step}/{object}_{kinematic}.pdf
예: /Users/user/overlay/TTbar_Signal/step5/Lep1_pT.pdf
```
파일 명은 object에 따라 다음 규칙을 따릅니다: 

| Object | Pattern |
|--------|------------|
| **Num_PV** | `h_Num_PV_{step_num}.{file_format}` |
| **Num_Jets/bJets** | `h_Num_{Jets/bJets}_{step_num}.{file_format}` |
| **Mass** | `h_DiLep_Mass_{step_num}.{file_format}` |
| **Lep, Jet, bJet, MET** | `h_{obj_name}_{kinematics}_{step_num}.{file_format}` |
| **Observable** | `h_Reco_CPO{obj_num}_ReRange.{file_format}` |
| **Initial stage** | `h_Num_PV_afterTrigger.{file_format}` |
| **After Top Reco** | `h_{obj_name}_{kinematics}.{file_format}` | 

만약 event selection이 'Initial stage'나 'After Top Reco'라면 별도의 형식을 따라 Plot을 찾아서 삽입합니다.

## 삽입 위치
선택된 object에 따라 달리 삽입됩니다.
- **PV, Mass**: 모든 event selection의 plot들을 하나의 슬라이드에 삽입합니다.
삽입위치는 event selection마다 고정되어 있으며, 몇몇개의 event selection만 선택했다면 순차적으로 삽입됩니다.
- **Num_Jets/bJets, MET, Observable, Nu와 AnNu**:
1x2 배열로 삽입됩니다.
- **Lep, Jet, bJet**: leading, subleading 모두 선택시 2x3 배열로 삽입됩니다.
만약 둘 중 하나만 선택한 경우 1x3 배열로 삽입됩니다.
- **bJet, AnbJet, Top, AnTop**: matter, antimatter 모두 선택시 2x4 배열로 삽입됩니다.
만약 둘 중 하나만 선택한 경우 1x4 배열로 삽입됩니다.
