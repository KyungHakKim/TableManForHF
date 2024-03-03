Attribute VB_Name = "basCONST"
Option Explicit

Public Const C_LOGIN_CONFIG_HEIGHT = 6285
Public Const C_LOGIN_HEIGHT = 2880
Public Const GCONN_USE = False
Public gConn As ADODB.Connection


'**********************************************************************
'                      공통RGB 선언
'**********************************************************************
Public Const S_BLACK = &H0&             'RGB(0, 0, 0)
Public Const S_WHITE = &HFFFFFF         'RGB(255, 255, 255)
Public Const S_LIGHT_GRAY = &HC0C0C0    'RGB(192, 192, 192)
Public Const S_DARK_GRAY = &H808080     'RGB(128, 128, 128)
Public Const S_RED = &HFF&              'RGB(255, 0, 0)
Public Const S_DARK_RED = &H80&         'RGB(128, 0, 0)
Public Const S_GREEN = &HFF00&          'RGB(0, 255, 0)
Public Const S_DARK_GREEN = &H8000&     'RGB(0, 128, 0)
Public Const S_BLUE = &HFF0000          'RGB(0, 0, 255)
Public Const S_DARK_BLUE = &H800000     'RGB(0, 0, 128)
Public Const S_MAGENTA = &HFF00FF       'RGB(255, 0, 255)
Public Const S_DARK_MAGENTA = &H800080  'RGB(128, 0, 128)
Public Const S_CYAN = &HFFFF00          'RGB(0, 255, 255)
Public Const S_DARK_CYAN = &H808000     'RGB(0, 128, 128)
Public Const S_YELLOW = &HFFFF&         'RGB(255, 255, 0)
Public Const S_BROWN = &H8080&          'RGB(128, 128, 0)

Public Const S_TOT1 = &HC0FFFF          '연한 노란색
Public Const S_TOT2 = &HC0C0&           '겨자색?, 황토색?

'**********************************************************************
'Spread 편집관련 COLOR 및 SYMBOL 선언
'**********************************************************************
Public Const C_ADD_COLOR = &HFCDD94
Public Const C_DEL_COLOR = &HD3C6FB
Public Const C_MOD_COLOR = &HBBFDFB
Public Const C_ORG_COLOR = &H80000005

Public Const C_ADD_SYMBOL = ">>"
Public Const C_DEL_SYMBOL = "X"
Public Const C_MOD_SYMBOL = ">"

Public Const C_ASC = &HE6CCD3
Public Const C_DESC = &HD2F2D8
Public Const C_NONE = &H8000000E


'**********************************************************
'SPREAD 관련 상수
'**********************************************************

' Action property settings
Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Public Const SS_ACTION_SMARTPRINT = 32

' Appearance property settings
Public Const SS_APPEARANCE_FLAT = 0
Public Const SS_APPEARANCE_3D = 1
Public Const SS_APPEARANCE_3DWITHBORDER = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1
Public Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY = 2
Public Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY = 3

' ButtonDrawMode property settings
Public Const SS_BDM_ALWAYS = 0
Public Const SS_BDM_CURRENT_CELL = 1
Public Const SS_BDM_CURRENT_COLUMN = 2
Public Const SS_BDM_CURRENT_ROW = 4
Public Const SS_BDM_ALWAYS_BUTTON = 8
Public Const SS_BDM_ALWAYS_COMBO = 16

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11

' ClipboardOptions property settings
Public Const SS_CLIP_NOHEADERS = 0
Public Const SS_CLIP_COPYROWHEADERS = 1
Public Const SS_CLIP_PASTEROWHEADERS = 2
Public Const SS_CLIP_COPYCOLHEADERS = 4
Public Const SS_CLIP_PASTECOLHEADERS = 8
Public Const SS_CLIP_COPYPASTEALLHEADERS = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7
Public Const SS_CURSOR_TYPE_DRAGDROPAREA = 8
Public Const SS_CURSOR_TYPE_DRAGDROP = 9

' DAutoSize property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' EditEnterAction property settings
Public Const SS_CELL_EDITMODE_EXIT_NONE = 0
Public Const SS_CELL_EDITMODE_EXIT_UP = 1
Public Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Public Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Public Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Public Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Public Const SS_CELL_EDITMODE_EXIT_SAME = 7
Public Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' Position property settings
Public Const SS_POSITION_UPPER_LEFT = 0
Public Const SS_POSITION_UPPER_CENTER = 1
Public Const SS_POSITION_UPPER_RIGHT = 2
Public Const SS_POSITION_CENTER_LEFT = 3
Public Const SS_POSITION_CENTER_CENTER = 4
Public Const SS_POSITION_CENTER_RIGHT = 5
Public Const SS_POSITION_BOTTOM_LEFT = 6
Public Const SS_POSITION_BOTTOM_CENTER = 7
Public Const SS_POSITION_BOTTOM_RIGHT = 8

' PrintOrientation property settings
Public Const SS_PRINTORIENT_DEFAULT = 0
Public Const SS_PRINTORIENT_PORTRAIT = 1
Public Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Public Const SS_PRINT_ALL = 0
Public Const SS_PRINT_CELL_RANGE = 1
Public Const SS_PRINT_CURRENT_PAGE = 2
Public Const SS_PRINT_PAGE_RANGE = 3

' ScrollBars property settings
Public Const SS_SCROLLBAR_NONE = 0
Public Const SS_SCROLLBAR_H_ONLY = 1
Public Const SS_SCROLLBAR_V_ONLY = 2
Public Const SS_SCROLLBAR_BOTH = 3

' ScrollBarTrack property settings
Public Const SS_SCROLLBARTRACK_OFF = 0
Public Const SS_SCROLLBARTRACK_VERTICAL = 1
Public Const SS_SCROLLBARTRACK_HORIZONTAL = 2
Public Const SS_SCROLLBARTRACK_BOTH = 3

' SelBackColor property settings
Public Const SPREAD_COLOR_NONE = &H8000000B

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' SortBy property settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' TextTip property settings
Public Const SS_TEXTTIP_OFF = 0
Public Const SS_TEXTTIP_FIXED = 1
Public Const SS_TEXTTIP_FLOATING = 2
Public Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Public Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

' TypeButtonAlign property settings
Public Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Public Const SS_CELL_BUTTON_ALIGN_TOP = 1
Public Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Public Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' TypeButtonType property settings
Public Const SS_CELL_BUTTON_NORMAL = 0
Public Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' TypeCheckType property settings
Public Const SS_CHECKBOX_NORMAL = 0
Public Const SS_CHECKBOX_THREE_STATE = 1

' TypeDateFormat property settings
Public Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Public Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Public Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Public Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Public Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Public Const SS_CELL_EDIT_CASE_NO_CASE = 1
Public Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Public Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Public Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeHAlign property settings
Public Const SS_CELL_H_ALIGN_LEFT = 0
Public Const SS_CELL_H_ALIGN_RIGHT = 1
Public Const SS_CELL_H_ALIGN_CENTER = 2

' TypeTextAlignVert property settings
Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Public Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Public Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Public Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Public Const SS_CELL_TIME_24_HOUR_CLOCK = 1

' TypeVAlign property settings
Public Const SS_CELL_V_ALIGN_TOP = 0
Public Const SS_CELL_V_ALIGN_BOTTOM = 1
Public Const SS_CELL_V_ALIGN_VCENTER = 2

' UnitType property settings
Public Const SS_CELL_UNIT_NORMAL = 0
Public Const SS_CELL_UNIT_VGA = 1
Public Const SS_CELL_UNIT_TWIPS = 2

' UserResize property settings
Public Const SS_USER_RESIZE_COL = 1
Public Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Public Const SS_USER_RESIZE_DEFAULT = 0
Public Const SS_USER_RESIZE_ON = 1
Public Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Public Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' ActionKey method settings
Public Const SS_KBA_CLEAR = 0
Public Const SS_KBA_CURRENT = 1
Public Const SS_KBA_POPUP = 2

' AddCustomFunctionExt method Flags parameter settings
Public Const SS_CUSTFUNC_WANTCELLREF = 1
Public Const SS_CUSTFUNC_WANTRANGEREF = 2

' CFGetParamInfo method Type parameter settings
Public Const SS_VALUE_TYPE_LONG = 0
Public Const SS_VALUE_TYPE_DOUBLE = 1
Public Const SS_VALUE_TYPE_STR = 2
Public Const SS_VALUE_TYPE_CELL = 3
Public Const SS_VALUE_TYPE_RANGE = 4

' CFGetParamInfo method Status parameter settings
Public Const SS_VALUE_STATUS_OK = 0
Public Const SS_VALUE_STATUS_ERROR = 1
Public Const SS_VALUE_STATUS_EMPTY = 2

' GetRefStyle/SetRefStyle methods return values/parameter settings
Public Const SS_REFSTYLE_DEFAULT = 0
Public Const SS_REFSTYLE_A1 = 1
Public Const SS_REFSTYLE_R1C1 = 2

' PrintOptions method PageOrder parameter settings
Public Const SS_PAGEORDER_AUTO = 0
Public Const SS_PAGEORDER_DOWNTHENOVER = 1
Public Const SS_PAGEORDER_OVERTHENDOWN = 2

' TextTipFetch method MultiLine parameter settings
Public Const SS_TT_MULTILINE_SINGLE = 0
Public Const SS_TT_MULTILINE_MULTI = 1
Public Const SS_TT_MULTILINE_AUTO = 2

' *************************  PrintPreview Settings *************************

' GrayAreaMarginType property values
Public Const SPV_GRAYAREAMARGINTYPE_SCALED = 0
Public Const SPV_GRAYAREAMARGINTYPE_ACTUAL = 1

' MousePointer property values
Public Const SPV_MOUSEPOINTER_DEFAULT = 0
Public Const SPV_MOUSEPOINTER_ARROW = 1
Public Const SPV_MOUSEPOINTER_CROSS = 2
Public Const SPV_MOUSEPOINTER_I_BEAM = 3
Public Const SPV_MOUSEPOINTER_ICON = 4
Public Const SPV_MOUSEPOINTER_SIZE = 5
Public Const SPV_MOUSEPOINTER_SIZE_NE_SW = 6
Public Const SPV_MOUSEPOINTER_SIZE_N_S = 7
Public Const SPV_MOUSEPOINTER_SIZE_NW_SE = 8
Public Const SPV_MOUSEPOINTER_SIZE_W_E = 9
Public Const SPV_MOUSEPOINTER_UP_ARROW = 10
Public Const SPV_MOUSEPOINTER_HOURGLASS = 11
Public Const SPV_MOUSEPOINTER_NO_DROP = 12

' PageViewType property values
Public Const SPV_PAGEVIEWTYPE_WHOLE_PAGE = 0
Public Const SPV_PAGEVIEWTYPE_NORMAL_SIZE = 1
Public Const SPV_PAGEVIEWTYPE_PERCENTAGE = 2
Public Const SPV_PAGEVIEWTYPE_PAGE_WIDTH = 3
Public Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT = 4
Public Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES = 5

' ScrollBarH property values
Public Const SPV_SCROLLBARH_SHOW = 0
Public Const SPV_SCROLLBARH_AUTO = 1
Public Const SPV_SCROLLBARH_HIDE = 2

' ScrollBarV property values
Public Const SPV_SCROLLBARV_SHOW = 0
Public Const SPV_SCROLLBARV_AUTO = 1
Public Const SPV_SCROLLBARV_HIDE = 2

' ZoomState property values
Public Const SPV_ZOOMSTATE_INDETERMINATE = 0
Public Const SPV_ZOOMSTATE_IN = 1
Public Const SPV_ZOOMSTATE_OUT = 2
Public Const SPV_ZOOMSTATE_SWITCH = 3

' Spread의 Cell별 Const 변수
Public Const SS_HEADER = 0
Public Const SSHD_ADD = ">"                ' Sheet에 add 상태로
Public Const SSHD_UPDATE = ">>"            ' Sheet에 update상태로
Public Const SSHD_DELETE = "X"             ' Sheet에 delete상태로
Public Const SSCELL_ADD = &HC0FFC0         ' 추가row의 BackColor
Public Const SSCELL_UPDATE = &HC0FFFF      ' 수정row의 BackColor
Public Const SSCELL_DELETE = &HC0C0FF      ' 삭제row의 BackColor
Public Const SSCELL_DEFAULT = &HFFFFFF     ' default ow의 BackColor
Public Const SSCELL_BORDER = &HC0C0C0      ' Border Color



