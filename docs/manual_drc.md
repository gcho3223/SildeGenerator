# DRC 모드 사용법

## DRC 모드
DRC Analysis에서 생성된 Plot들을 슬라이드에 자동으로 삽입하는 모드입니다. 
PID의 최종단계의 Energy deposit distribution과 Energy resolution & linearity plot을 삽입할 수 있습니다.
### **주의: 이 모드는 디렉토리 구조 및 파일 이름이 나한테 맞춰져있음니다**

## 1. Option
### 1) Particles
- EM
- Pion
- Proton
- Kaon

### 2) Cases
3가지 case가 있습니다.
- Normal (일반, rotation, tilting, interaction target 없는 상태)
- Rotation & Tilting (각각 +1.5 degree, +1.0 degree로 검출기를 회전시키고 기울인 상태)
- Interaction target (검출기 앞에 Interaction target을 설치한 상태)

### 3) File Format
- PDF, PNG 파일을 지원합니다.
- JPG, JPEG 등 다른 사진파일들은 필요시 업데이트 예정

## 2. Channel & Energy Points
### 1) Channel 선택
- **C**: Cerenkov channel
- **S**: Scintilation channe
- **DRcor**: Dual-readout correction 적용
- **LC+ATT correction**
    - S channel에는 Leakage counter, Attenuation correction이 적용되므로 각각에 대한 옵션이 있습니다.
    - 따라서 S, DRcor channel에는 correction 옵션이 있습니다.
        - None: 삽입하지 않음
        - S/DRcor: LC, ATT correction 미적용
        - LCcor: Leakage counter correction만 적용
        - ATTcor: Attenueation correction만 적용
        - LC+ATTcor: Leakage counter + Attenuation correction 모두 적용

### 2) Energy Points 선택
**SPS(H8) 기준 Energy range**
- 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120 GeV를 지원합니다.
- 삽입할 energy를 선택합니다.

**Low Energy**
- **KEK** 빔테스트에 쓰인 low energy range를 지원합니다.
- "Low Energy" 체크박스 활성화
- 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0 GeV를 지원합니다.

### 3) Resolution & Linearity
- **With noise term**: Noise term을 포함한 Energy resolution plot
- **Without noise term**: Noise term을 포함하지 않는 Energy resolution plot
- **Linearity**
※ Resolution과 Linearity는 별도 슬라이드로 생성됩니다

## 3. PID
- Muon counter까지 PID를 거친 최종단계의 plot만 지원합니다.

## 4. Preview
Energy deposit plot의 삽입 위치를 미리 보여줍니다.
Energy가 작은 순서대로 배치되며 이 위치는 드래그 앤 드랍으로 옮길 수 있습니다.

## 파일 경로 규칙
DRC 모드에서는 다음 경로 규칙을 따릅니다:
```
{base_dir}/{program}/c_{particle}_{channel}_{energy}.pdf
예: /Users/user/TB2025v2/proton_Rot/c_proton_C_energy20_run20.pdf

Resolution/Linearity:
{base_dir}/{program}/c_{particle}Resol_M5-T2.pdf
```