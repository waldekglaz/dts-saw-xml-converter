import pandas as pd
import glob
import os

def process_file(excel_file):
    print(f"Processing file: {excel_file}")
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_file, header=0)
    except Exception as e:
        print(f"Error reading file '{excel_file}': {e}")
        return

    # Check if we have enough columns
    if len(df.columns) < 5:
        print(f"Skipping '{excel_file}': Must have at least 5 columns (Width, Length, Qty, Grain, Description)")
        return

    # Prepare XML content
    xml_parts = []
    xml_parts.append("<?xml version='1.0' encoding='UTF-8'?>")
    n_parts = len(df)
    xml_parts.append(f"<CutList NParts='{n_parts}' NBoards='1'>")

    for index, row in df.iterrows():
        try:
            width = float(row.iloc[0])
            length = float(row.iloc[1])
            qty = int(row.iloc[2]) if not pd.isna(row.iloc[2]) else 0
            
            raw_grain = row.iloc[3]
            grain = 0
            if not pd.isna(raw_grain):
                s_grain = str(raw_grain).strip().lower()
                if s_grain in ['1', 'yes', 'y', 'true']:
                    grain = 1
                elif s_grain in ['0', 'no', 'n', 'false']:
                    grain = 0
                else:
                    try:
                        grain = int(float(raw_grain))
                    except:
                        grain = 0

            desc = str(row.iloc[4]) if not pd.isna(row.iloc[4]) else ""
            
            # Check for optional 6th column (Secondary Description)
            sec_desc = ""
            if len(row) > 5:
                 val = row.iloc[5]
                 if not pd.isna(val):
                     sec_desc = str(val)

        except Exception as e:
            print(f"Skipping row {index+2} in '{excel_file}' due to error: {e}")
            continue

        xml_line = f" <Part id='P{index+1}' L='{length:.2f}' W='{width:.2f}' qMin='{qty}' Grain='{grain}' IDesc='{desc}'"
        if sec_desc:
            xml_line += f" IIDesc='{sec_desc}'"
        xml_line += "/>"
        
        xml_parts.append(xml_line)

    # Static blocks
    static_content = """ <Board id='B1' L='2440.00' W='1220.00' Thickness='18.00' TTrim='0.00' LTrim='0.00' MatNo='4' MatCode='MDF' Qty='55555' Stock='1'/>
 <Param Algo='2' LR='1' SR='1' SHC='1'>
  <OptiParam ver='6'/>
  <OverParam/>
  <DropParam/>
  <StackParam MaxStack='100'/>
  <BundleParam MinBundle='1' MaxBundle='0'/>
  <ZCutParam MaxStages='3'/>
  <LRParam ZCutsAllowed='1' />
  <SRParam ZCutsAllowed='1' />
  <SHCParam>
   <ShortHeadCut ZCutsAllowed='1' />
   <ShortMainBody ZCutsAllowed='1' />
  </SHCParam>
  <LHCParam>
   <LongHeadCut />
   <LongMainBody />
  </LHCParam>
 </Param>
 <SimulParam Model='0' TimeCost='0'>
  <Blades Rip='4.20' Cross='4.20' HC='4.20' ZC='4.20'/>
  <Trims MinDivRip='0.90' MaxDivRip='40.00' OptRip='10.00' MinDivCross='0.90' MaxDivCross='40.00' OptCross='10.00' HCTrim='20.00' HCTrimMin='0.90'/>
  <Clamps Nr='0' Dim='0.00' Tol='0.00' ClampsTime='0'/>
  <TMan />
  <FirstAxis />
  <SecondAxis />
  <Shuttle />
  <LiftTable />
  <Vacuum />
  <TurningTable />
 </SimulParam>
</CutList>"""
    
    xml_parts.append(static_content)

    # Output file
    output_filename = os.path.splitext(excel_file)[0] + ".xml"
    with open(output_filename, "w", encoding='utf-8') as f:
        f.write("\n".join(xml_parts))
    
    print(f"Successfully created {output_filename}")

def main():
    # Find all xlsx files
    files = glob.glob('*.xlsx')
    if not files:
        print("No Excel files found in the current directory.")
        return

    print(f"Found {len(files)} Excel file(s).")
    for excel_file in files:
        process_file(excel_file)

if __name__ == "__main__":
    main()
