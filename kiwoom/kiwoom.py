from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("kiwoom입니다.")
        ######## eventloop 모음 #########
        self.login_event_loop = None
        self.account_num = None
        #################################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info() # 예수금 가져오는 부분
        ####### 테스트 ##############
        self.test()
        #############################

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec()

    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))

        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")

        self.account_num = account_list.split(';')[0]
        print("나의 보유 계좌번호 %s" % self.account_num)

    def detail_account_info(self):
        print("예수금을 요청하는 부분")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "0000")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", "2000")

    ####test#################################################################
    def test(self):
        print("test")

        self.dynamicCall("SetInputValue(String, String)", "시작일자", "20200401")
        self.dynamicCall("SetInputValue(String, String)", "종료일자", "20200404")
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "1")
        self.dynamicCall("SetInputValue(String, String)", "시장구분", "001")
        self.dynamicCall("SetInputValue(String, String)", "투자자구분", "8000")
        self.dynamicCall("CommRqData(String, String, int, String)", "투자자별일별매매종목요청", "OPT10058", "0", "2000")

    #########################################################################



    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr요청을 받는 구역이다! 슬롯이다!
        :param sScrNo: 스크린 번호
        :param sRQName: 내가 요청했을떄 지은 이름
        :param sTrCode: 요청id, tr코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print("예수금 : %s" % int(deposit))
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액 : %s" % int(ok_deposit))


        ####test#################################################################
        if sRQName == "투자자별일별매매종목요청":
            test = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "순매도수량")
            print("test: %s" % test)
        #########################################################################






