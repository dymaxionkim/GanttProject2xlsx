# GanttProject2xlsx
_Automake Spreadsheet Report from GanttProject_



## Why

* 직장상사들은 MS엑셀 포멧 문서를 엄청나게 좋아하는 경향이 강하다.
* 직장상사들은 1페이지에 업무현황이 요약된 간략한 보고서를 원한다.
* 직장상사들은 매일매일 보고서를 받기를 원한다.
* 그런데 매일매일 보고서를 일일이 쓰는 것은 고문에 가까운 일이다.
* GanttProject에서 직접 만들어주는 간트챠트는 예쁘지도 않고 보기가 불편하다.  다른 프로젝트 관리도구들도 대체로 별로 좋지가 않다.  직장상사들은 이런 디자인을 싫어한다.
* 스케쥴 관리하기 위한 자료 입력에는 GanttProject가 상당히 괜챦다.  너무 복잡하지도 않고 심플하다.
* 내용입력과 편집은 GanttProject로 심플하게 하고, 직장상사가 원하는 형태의 보고서를 자동으로 만들어내서 곧바로 출력해서 갖다 드리면 칼퇴근 가능.



## Prerequisites

* Linux
* [GanttProject](https://www.ganttproject.biz/) ([setting_01](setting_01.png), [setting_02](setting_02.png))
* Python3 with Pandas, datetime, openpyxl  (Fix execution path in `gantt.sh`)
* Libreoffice (to export pdf)



## How to use

* Edit `gantt.gan` with GanttProject
* Run `gantt.sh`
* Open `Report.xlsx`, `Report.pdf` and hit `Ctrl+P` to print

