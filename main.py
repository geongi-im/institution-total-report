import requests
from dotenv import load_dotenv
import os
import sys
import json
from datetime import datetime
import pandas as pd
import imgkit
from utils.api_util import ApiUtil, ApiError
from utils.telegram_util import TelegramUtil
from utils.logger_util import LoggerUtil
import holidays
import pykrx.stock as stock

load_dotenv()
  
# 특정 종목코드가 어느 시장에 속하는지 확인
def checkMarket(ticker):
    if ticker in kospi_tickers:
        return "KOSPI"
    elif ticker in kosdaq_tickers:
        return "KOSDAQ"
    else:
        return "Not Found"

def isTodayHoliday():
    kr_holidays = holidays.KR()
    today = datetime.today().date()
    return today in kr_holidays

class InstitutionTotalReport:
    def __init__(self):
        self.url_base = os.getenv("KIS_URL_BASE")
        self.app_key = os.getenv("KIS_APP_KEY")
        self.app_secret = os.getenv("KIS_APP_SECRET")
        self.token_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'token.json')
        self.img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
        self.wkhtmltoimage_path = os.getenv('WKHTMLTOIMAGE_PATH')
        self.logger = LoggerUtil().get_logger()
        
        # img 디렉토리가 없으면 생성
        if not os.path.exists(self.img_dir):
            os.makedirs(self.img_dir)
            self.logger.info(f"이미지 디렉토리 생성: {self.img_dir}")

    def load_token(self):
        """토큰 파일에서 저장된 토큰 정보를 로드"""
        if not os.path.exists(self.token_file):
            self.logger.debug("토큰 파일이 존재하지 않습니다.")
            return None
            
        with open(self.token_file, 'r') as f:
            data = json.load(f)
            
        # 만료 시간 확인 (둘 중 하나라도 만료되면 새로운 토큰 발급)
        now = datetime.now()
        expires_at = datetime.strptime(data['access_token_token_expired'], "%Y-%m-%d %H:%M:%S")
        
        if expires_at <= now:
            self.logger.debug("토큰이 만료되었습니다.")
            return None
            
        self.logger.debug("유효한 토큰을 로드했습니다.")
        return data['access_token']

    def save_token(self, token_info):
        """토큰 정보를 파일에 저장
        token_info: API 응답의 토큰 정보 (access_token, expires_in, access_token_token_expired 포함)
        """
        data = {
            'access_token': token_info['access_token'],
            'expires_in': token_info['expires_in'],  # 유효기간(초)
            'access_token_token_expired': token_info['access_token_token_expired']  # 만료일시
        }
        
        with open(self.token_file, 'w') as f:
            json.dump(data, f)
            
        self.logger.debug(f"토큰 정보를 저장했습니다. 만료일시: {token_info['access_token_token_expired']}")

    def check_env_variables(self):
        """필수 환경변수 체크"""
        required_vars = ['KIS_APP_KEY', 'KIS_APP_SECRET', 'KIS_URL_BASE']
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        
        if missing_vars:
            error_msg = f"다음 환경변수가 설정되지 않았습니다: {', '.join(missing_vars)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)

    def get_token(self):
        """토큰 조회 또는 새로 발급"""
        # 환경변수 체크
        self.check_env_variables()
        
        # 저장된 토큰이 있는지 확인
        token = self.load_token()
        if token:
            return token

        # 새로운 토큰 발급
        self.logger.info("새로운 토큰 발급 시작")
        headers = {"content-type":"application/json"}
        body = {
            "grant_type":"client_credentials",
            "appkey": self.app_key,
            "appsecret": self.app_secret
        }
        PATH = "oauth2/tokenP"
        URL = f"{self.url_base}/{PATH}"
        
        res = requests.post(URL, headers=headers, data=json.dumps(body))
        
        if res.status_code != 200:
            error_msg = "토큰 발급 실패"
            self.logger.error(error_msg)
            raise Exception(error_msg)
            
        token_info = res.json()
        self.save_token(token_info)
        self.logger.info("토큰 발급 성공")
        
        return token_info['access_token'] 
    
    def get_institution_total_report(self):
        token = self.get_token()
        if not token:
            error_msg = "토큰 발급 실패"
            self.logger.error(error_msg)
            raise Exception(error_msg)
            
        # API 엔드포인트 설정
        PATH = "uapi/domestic-stock/v1/quotations/foreign-institution-total"
        # PATH = "uapi/domestic-stock/v1/quotations/inquire-price"
        URL = f"{self.url_base}/{PATH}"

        self.logger.info("기관 순매수 데이터 조회 시작")
        
        # 요청 헤더 설정
        headers = {
            "Content-Type": "application/json; charset=utf-8", 
            "authorization": f"Bearer {token}",
            "appKey": self.app_key,
            "appSecret": self.app_secret,
            "tr_id": "FHPTJ04400000",
        }

        # 요청 파라미터 설정
        params = {
            "FID_COND_MRKT_DIV_CODE": "V",
            "FID_COND_SCR_DIV_CODE": "16449",
            "FID_INPUT_ISCD": "0000",
            "FID_DIV_CLS_CODE": "0", #0:수량정렬, 1:금액정렬
            "FID_RANK_SORT_CLS_CODE": "0", #0:순매수상위, 1:순매도상위
            "FID_ETC_CLS_CODE": "2" #0:전체, 1:외국인, 2:기관계, 3:기타
        }

        # API 호출
        res = requests.get(URL, headers=headers, params=params)
        if res.status_code == 200 and res.json()["rt_cd"] == "0":
            result = res.json()["output"]
            self.logger.info(f"기관 순매수 데이터 조회 성공: {len(result)}개 종목")
            return result
        else:
            error_msg = f"API 호출 실패: {res.json()['msg_cd']}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
    
    def get_stock_price(self, stock_code, start_date=None, end_date=None):
        """특정 종목의 주가 정보를 조회하는 함수
        
        Args:
            stock_code (str): 종목코드 (6자리)
            start_date (str, optional): 조회 시작일 (YYYYMMDD 형식). 기본값은 100일 전
            end_date (str, optional): 조회 종료일 (YYYYMMDD 형식). 기본값은 현재일
            
        Returns:
            pandas.DataFrame: 주가 데이터
        """
        token = self.get_token()
        if not token:
            error_msg = "토큰 발급 실패"
            self.logger.error(error_msg)
            raise Exception(error_msg)
            
        # 날짜 파라미터 설정
        if start_date is None:
            start_date = (datetime.now() - pd.Timedelta(days=100)).strftime("%Y%m%d")
        if end_date is None:
            end_date = datetime.today().strftime("%Y%m%d")
            
        self.logger.debug(f"주가 조회 시작 - 종목코드: {stock_code}, 조회기간: {start_date} ~ {end_date}")
            
        # API 엔드포인트 설정
        PATH = "uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
        URL = f"{self.url_base}/{PATH}"
        
        # 요청 헤더 설정
        headers = {
            "Content-Type": "application/json; charset=utf-8", 
            "authorization": f"Bearer {token}",
            "appKey": self.app_key,
            "appSecret": self.app_secret,
            "tr_id": "FHKST03010100",  # 국내주식기간별시세
        }
        
        # 요청 파라미터 설정
        params = {
            "FID_COND_MRKT_DIV_CODE": "J",  # 시장 분류 코드 J:주식/ETF/ETN
            "FID_INPUT_ISCD": stock_code,    # 종목번호 (6자리)
            "FID_INPUT_DATE_1": start_date,  # 조회 시작일자
            "FID_INPUT_DATE_2": end_date,    # 조회 종료일자
            "FID_PERIOD_DIV_CODE": "D",      # 기간분류코드 D:일봉
            "FID_ORG_ADJ_PRC": "1"           # 수정주가 여부 (0:수정주가, 1:원주가)
        }
        
        # API 호출
        res = requests.get(URL, headers=headers, params=params)
        
        if res.status_code == 200 and res.json()["rt_cd"] == "0":
            # 주가 데이터를 DataFrame으로 변환
            data = res.json()["output2"]  # output2에 시계열 데이터가 포함됨
            df = pd.DataFrame(data)
            
            # 컬럼 이름 변경 및 데이터 타입 변환
            rename_cols = {
                'stck_bsop_date': '날짜',
                'stck_oprc': '시가',
                'stck_hgpr': '고가',
                'stck_lwpr': '저가',
                'stck_clpr': '종가',
                'acml_vol': '거래량',
                'acml_tr_pbmn': '거래대금',
                'flng_cls_code': '등락구분',
                'prtt_rate': '등락률',
                'mod_yn': '분할여부',
                'prdy_vrss': '전일대비'
            }
            
            # 컬럼 선택 및 이름 변경
            cols_to_use = list(rename_cols.keys())
            df = df[cols_to_use].rename(columns=rename_cols)
            
            # 데이터 타입 변환
            numeric_cols = ['시가', '고가', '저가', '종가', '거래량', '거래대금', '등락률', '전일대비']
            for col in numeric_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 날짜 형식 변환
            df['날짜'] = pd.to_datetime(df['날짜'], format='%Y%m%d')
            
            # 날짜 기준 내림차순 정렬
            df = df.sort_values(by='날짜', ascending=False).reset_index(drop=True)
            
            self.logger.debug(f"주가 조회 완료 - 종목코드: {stock_code}, 데이터 수: {len(df)}")
            return df
        else:
            error_msg = res.json().get("msg_cd", "알 수 없는 오류")
            self.logger.error(f"주가 조회 실패 - 종목코드: {stock_code}, 오류: {error_msg}")
            raise Exception(f"API 호출 실패: {error_msg}")

    def get_domestic_index(self, market_code="KOSPI", date=None, period="D"):
        """국내 주요 지수 데이터를 조회하는 함수
        
        Args:
            market_code (str, optional): 시장 코드. KOSPI 또는 KOSDAQ. 기본값은 KOSPI
            date (str, optional): 조회일자 (YYYYMMDD 형식). 기본값은 현재일
            period (str, optional): 기간분류코드 D:일, W:주, M:월, Y:년. 기본값은 일봉(D)
            
        Returns:
            pandas.DataFrame: 지수 데이터
        """
        token = self.get_token()
        if not token:
            raise Exception("토큰 발급 실패")
                        
        # 날짜 파라미터 설정
        if date is None:
            date = datetime.today().strftime("%Y%m%d")

        if market_code == "KOSPI":
            market_code = "0001"
        elif market_code == "KOSDAQ":
            market_code = "1001"
            
        # API 엔드포인트 설정
        PATH = "uapi/domestic-stock/v1/quotations/inquire-index-daily-price"
        URL = f"{self.url_base}/{PATH}"
        
        # 요청 헤더 설정
        headers = {
            "Content-Type": "application/json; charset=utf-8", 
            "authorization": f"Bearer {token}",
            "appKey": self.app_key,
            "appSecret": self.app_secret,
            "tr_id": "FHPUP02120000",  # 국내업종 일자별지수[v1_국내주식-065]
        }
        
        # 요청 파라미터 설정
        params = {
            "FID_COND_MRKT_DIV_CODE": "U",   # 시장구분코드 (업종 U)
            "FID_INPUT_ISCD": market_code,   # 시장 코드 (코스피(0001), 코스닥(1001), 코스피200(2001))
            "FID_INPUT_DATE_1": date,  # 입력 날짜(ex. 20240223)
            "FID_PERIOD_DIV_CODE": period    # 기간분류코드 D:일, W:주, M:월
        }
        
        # API 호출
        res = requests.get(URL, headers=headers, params=params)
        
        if res.status_code == 200 and res.json()["rt_cd"] == "0":
            # 지수 데이터를 DataFrame으로 변환
            data = res.json()["output2"]
            df = pd.DataFrame(data)
            
            # 컬럼 이름 변경 및 데이터 타입 변환
            rename_cols = {
                'stck_bsop_date': '날짜',
                'bstp_nmix_prpr': '종가',
                'bstp_nmix_oprc': '시가',
                'bstp_nmix_hgpr': '고가',
                'bstp_nmix_lwpr': '저가',
                'acml_vol': '거래량',
                'bstp_nmix_prdy_vrss': '전일대비',
                'prdy_vrss_sign': '등락구분',
                'bstp_nmix_prdy_ctrt': '등락률'
            }
            
            # 컬럼 선택 및 이름 변경
            cols_to_use = list(set(rename_cols.keys()) & set(df.columns))
            df = df[cols_to_use].rename(columns={col: rename_cols[col] for col in cols_to_use})
            
            # 지수명 컬럼 추가
            df['지수명'] = market_code
            
            # 데이터 타입 변환
            numeric_cols = ['종가', '시가', '고가', '저가', '거래량', '전일대비', '등락률']
            numeric_cols = [col for col in numeric_cols if col in df.columns]
            for col in numeric_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 날짜 형식 변환
            if '날짜' in df.columns:
                df['날짜'] = pd.to_datetime(df['날짜'], format='%Y%m%d')
                # 날짜 기준 내림차순 정렬
                df = df.sort_values(by='날짜', ascending=False).reset_index(drop=True)
            
            return df
        else:
            error_message = res.json().get("msg_cd", "알 수 없는 오류")
            raise Exception(f"API 호출 실패: {error_message}")

    def add_historical_price_change(self, filtered_data, reference_date):
        """기관 순매수 데이터에 과거 가격 대비 현재 가격 등락률을 추가하는 함수
        
        Args:
            filtered_data (list): 기관 순매수 데이터 리스트
            reference_date (str): 과거 가격 조회 기준일(YYYYMMDD 형식)
            
        Returns:
            list: 등락률이 추가된 기관 순매수 데이터 리스트
        """
        result = []
        
        self.logger.info(f"총 {len(filtered_data)}개 종목의 과거 가격 조회 시작 - 기준일: {reference_date}")
        
        for idx, item in enumerate(filtered_data):
            # 종목코드 추출
            stock_code = item['mksc_shrn_iscd']
            stock_name = item['hts_kor_isnm']
            # 현재가 추출 (문자열을 정수로 변환)
            current_price = int(item['stck_prpr'])
            
            self.logger.debug(f"{idx+1}/{len(filtered_data)} - {stock_name}({stock_code}) 과거 가격 조회")
            
            try:
                # 과거 가격 조회
                historical_data = self.get_stock_price(stock_code, start_date=reference_date, end_date=reference_date)
                
                # 과거 데이터가 존재하는 경우
                if not historical_data.empty:
                    # 과거 종가 추출
                    historical_price = historical_data.iloc[0]['종가']
                    
                    # 등락률 계산 (백분율)
                    if historical_price > 0:
                        price_change_rate = ((current_price - historical_price) / historical_price) * 100
                    else:
                        price_change_rate = 0
                    
                    # 원본 데이터를 복사하고 등락률 추가
                    item_copy = item.copy()
                    item_copy['historical_price'] = int(historical_price)
                    item_copy['price_change_rate'] = round(price_change_rate, 2)
                    result.append(item_copy)
                    
                    self.logger.debug(f"{stock_name} - 현재가: {current_price}, 과거가: {int(historical_price)}, 등락률: {round(price_change_rate, 2)}%")
                else:
                    # 과거 데이터가 없는 경우 원본 데이터를 유지
                    item_copy = item.copy()
                    item_copy['historical_price'] = 0
                    item_copy['price_change_rate'] = 0
                    result.append(item_copy)
                    
                    self.logger.warning(f"{stock_name} - 과거 데이터 없음")
            except Exception as e:
                # 오류 발생 시 원본 데이터를 유지
                self.logger.error(f"오류: 종목 {stock_code} 과거 가격 조회 실패: {str(e)}")
                item_copy = item.copy()
                item_copy['historical_price'] = 0
                item_copy['price_change_rate'] = 0
                result.append(item_copy)
                
        self.logger.info(f"과거 가격 조회 및 등락률 계산 완료 - {len(result)}개 종목")
        return result

    def add_market_info_and_index_rate(self, enhanced_data, kospi_index_change_rate, kosdaq_index_change_rate):
        """기관 순매수 데이터에 시장 정보와 해당 시장 지수 등락률을 추가하는 함수
        
        Args:
            enhanced_data (list): 과거 가격 비교 등락률이 추가된 데이터 리스트
            kospi_index_change_rate (float): 코스피 지수 등락률
            kosdaq_index_change_rate (float): 코스닥 지수 등락률
            
        Returns:
            list: 시장 정보와 지수 등락률이 추가된 데이터 리스트
        """
        result = []
        
        self.logger.info(f"총 {len(enhanced_data)}개 종목의 시장 정보 조회 시작")
        self.logger.info(f"코스피 지수 등락률: {kospi_index_change_rate}%, 코스닥 지수 등락률: {kosdaq_index_change_rate}%")
        
        kospi_count = 0
        kosdaq_count = 0
        other_count = 0
        
        for idx, item in enumerate(enhanced_data):
            # 종목코드 추출
            stock_code = item['mksc_shrn_iscd']
            stock_name = item['hts_kor_isnm']
            
            # 시장 구분 확인
            market = checkMarket(stock_code)
            
            # 원본 데이터를 복사하고 시장 정보 및 지수 등락률 추가
            item_copy = item.copy()
            item_copy['market'] = market
            
            # 해당 시장의 지수 등락률 추가
            if market == "KOSPI":
                item_copy['index_change_rate'] = kospi_index_change_rate
                kospi_count += 1
            elif market == "KOSDAQ":
                item_copy['index_change_rate'] = kosdaq_index_change_rate
                kosdaq_count += 1
            else:
                item_copy['index_change_rate'] = 0
                other_count += 1
                self.logger.warning(f"{stock_name}({stock_code}) - 알 수 없는 시장")
                
            # 종목의 등락률과 시장 지수 등락률의 차이 계산
            # item_copy['outperform_rate'] = round(item_copy['price_change_rate'] - item_copy['index_change_rate'], 2)
            
            self.logger.debug(f"{idx+1}/{len(enhanced_data)} - {stock_name}({stock_code}): {market} 시장")
            
            result.append(item_copy)
            
        self.logger.info(f"시장 정보 추가 완료 - KOSPI: {kospi_count}개, KOSDAQ: {kosdaq_count}개, 기타: {other_count}개")
        return result

    def convert_to_dataframe(self, data, top_n=10):
        """API 응답 데이터를 DataFrame으로 변환"""
        if not data:
            self.logger.warning("데이터가 없어 DataFrame 변환 불가")
            return pd.DataFrame()
        
        self.logger.info(f"DataFrame 변환 시작 - 총 {len(data)}개 항목, top_n={top_n}")
        
        # 상위 N개만 필터링
        filtered_data = data[:top_n] if len(data) > top_n else data
        
        # 필요한 컬럼 추출
        df = pd.DataFrame(filtered_data)
        
        self.logger.debug(f"DataFrame 변환 - 컬럼: {list(df.columns)}")
        
        # 종목명과 종목코드 합치기 전에 별도 DataFrame 생성
        result_df = pd.DataFrame()
        
        # 종목명과 종목코드 합치기
        result_df['종목명'] = df.apply(lambda row: f"{row['hts_kor_isnm']} <span class='stock-code'>({row['mksc_shrn_iscd']})</span>", axis=1)
        result_df['현재가'] = df['stck_prpr'].astype(int).map('{:,}'.format)
        
        
        # 전일대비율에 색상 추가
        def format_rate(value):
            value_float = float(value)
            if value_float < 0:
                return f"<span class='negative'>{value_float:.2f}%</span>"
            elif value_float > 0:
                return f"<span class='positive'>{value_float:.2f}%</span>"
            else:
                return f"{value_float:.2f}%"
        
        # 시장등락률(30일)과 종목등락률(30일)을 하나로 합치기
        def format_compare_rates(row):
            market = row['market']
            market_rate = float(row['index_change_rate'])
            stock_rate = float(row['price_change_rate'])
            
            market_text = f"{market}: "
            if market_rate < 0:
                market_text += f"<span class='negative'>{market_rate:.2f}%</span>"
            elif market_rate > 0:
                market_text += f"<span class='positive'>{market_rate:.2f}%</span>"
            else:
                market_text += f"{market_rate:.2f}%"
                
            stock_text = "종목: "
            if stock_rate < 0:
                stock_text += f"<span class='negative'>{stock_rate:.2f}%</span>"
            elif stock_rate > 0:
                stock_text += f"<span class='positive'>{stock_rate:.2f}%</span>"
            else:
                stock_text += f"{stock_rate:.2f}%"
                
            return f"{market_text}<br>{stock_text}"
        
        # 시장대비등락률 컬럼 추가
        result_df['시장대비등락률'] = df.apply(format_compare_rates, axis=1)
        
        # result_df['전일대비율(%)'] = df['prdy_ctrt'].apply(format_rate)
        result_df['기관순매수량'] = df['orgn_ntby_qty'].astype(int).map('{:,}'.format)
        result_df['기관순매수금액'] = (df['orgn_ntby_tr_pbmn'].astype(float) / 100).round(2).map('{:,}'.format)  # 억원 단위로 변환
        
        self.logger.info(f"DataFrame 변환 완료 - 결과 컬럼: {list(result_df.columns)}")
        return result_df
    
    def save_df_as_image(self, df, file_name="institution_top_report"):
        """DataFrame을 이미지로 저장하고 파일 경로 반환"""
        if df.empty:
            self.logger.warning("DataFrame이 비어 있어 이미지를 생성할 수 없습니다.")
            return None
            
        self.logger.info(f"이미지 생성 시작 - 파일명: {file_name}")
            
        if not file_name.endswith('.png'):
            file_name = file_name + '.png'
            
        file_name, file_extension = os.path.splitext(file_name)
        current_date = datetime.now().strftime('%Y%m%d')
        new_file_path = os.path.join(self.img_dir, f"{file_name}_{current_date}{file_extension}")
        
        # 이전 파일 삭제
        removed_count = 0
        for old_file in os.listdir(self.img_dir):
            if old_file.startswith(file_name) and old_file.endswith(file_extension):
                try:
                    os.remove(os.path.join(self.img_dir, old_file))
                    removed_count += 1
                    self.logger.debug(f"기존 파일 삭제: {old_file}")
                except Exception as e:
                    self.logger.warning(f"파일 삭제 실패: {old_file} - {str(e)}")
                    
        self.logger.info(f"{removed_count}개의 기존 파일 삭제 완료")

        # 캡션 설정
        today_display = datetime.now().strftime('%Y-%m-%d')
        caption = f"{today_display} 기관 순매수 상위 TOP 10"
        
        self.logger.debug("HTML 생성 시작")

        # HTML 생성
        html_str = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet">
            <style>
                body {{
                    font-family: 'Noto Sans KR', sans-serif;
                    margin: 10px;
                    padding: 0;
                    max-width: 600px;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin: 10px auto;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                }}
                th, td {{
                    border: 1px solid #e0e0e0;
                    padding: 8px 10px;
                    text-align: center;
                }}
                th {{
                    background-color: #333333;
                    color: white;
                    font-weight: 700;
                    font-size: 13px;
                    white-space: nowrap;
                }}
                td {{
                    font-size: 12px;
                    font-weight: 500;
                }}
                td.stock-name {{
                    text-align: center;
                }}
                .stock-code {{
                    font-size: 10px;
                    color: #666;
                    display: block;
                    margin-top: 2px;
                }}
                tr:nth-child(even) td {{
                    background-color: #f9f9f9;
                }}
                tr:hover td {{
                    background-color: #f5f5f5;
                }}
                .caption {{
                    text-align: center;
                    font-size: 18px;
                    font-weight: 700;
                    margin: 15px 0;
                    color: #333333;
                }}
                .source {{
                    text-align: right;
                    font-size: 11px;
                    color: #666666;
                    margin-top: 10px;
                    font-weight: 400;
                }}
                .positive {{
                    color: #d32f2f;
                }}
                .negative {{
                    color: #1976d2;
                }}
            </style>
        </head>
        <body>
            <div class="caption">{caption}</div>
        '''

        # DataFrame을 HTML로 변환하고 종목명 열에 class 추가
        df_html = df.to_html(index=False, classes='styled-table', escape=False)

        # 시장대비등락률 헤더를 시장대비등락률<br>(30일기준)으로 변경
        df_html = df_html.replace('>시장대비등락률<', '>시장대비등락률<br>(30일기준)<')
        
        # 순매수금액 헤더를 순매수금액<br>(억원)으로 변경
        df_html = df_html.replace('>기관순매수금액<', '>기관순매수금액<br>(억원)<')
        
        # 종목명 열에 class 추가 (더 정확한 패턴으로 수정)
        import re
        
        # 헤더에서 첫 번째 <th>종목명</th> 패턴 찾기
        df_html = df_html.replace('<th>종목명</th>', '<th class="stock-name">종목명</th>')
        
        # 데이터 행에서 첫 번째 <td> 태그를 <td class="stock-name"> 으로 변경
        pattern = r'(<tr[^>]*>)\s*<td([^>]*)>'
        repl = r'\1<td class="stock-name"\2>'
        df_html = re.sub(pattern, repl, df_html)

        html_str += df_html
        html_str += '''
            <div class="source">※ 출처 : MQ(Money Quotient)</div>
        </body>
        </html>
        '''
        
        self.logger.debug("HTML 생성 완료")

        options = {
            'format': 'png',
            'encoding': "UTF-8",
            'quality': 100,
            'width': 600,
            'enable-local-file-access': None,
            'minimum-font-size': 10
        }

        try:
            if not self.wkhtmltoimage_path:
                error_message = "❌ 오류 발생\n\nWKHTMLTOIMAGE_PATH 환경변수가 설정되지 않았습니다."
                telegram = TelegramUtil()
                telegram.send_test_message(error_message)
                self.logger.error("WKHTMLTOIMAGE_PATH 환경변수가 설정되지 않았습니다.")
                raise ValueError("WKHTMLTOIMAGE_PATH 환경변수가 필요합니다.")
                
            config = imgkit.config(wkhtmltoimage=self.wkhtmltoimage_path)
            self.logger.info("이미지 생성 중...")
            imgkit.from_string(html_str, new_file_path, options=options, config=config)
            self.logger.info(f"새 파일 저장 완료: {new_file_path}")
            
            return new_file_path
            
        except Exception as e:
            error_message = f"❌ 오류 발생\n\n함수: save_df_as_image\n파일: {file_name}\n오류: {str(e)}"
            telegram = TelegramUtil()
            telegram.send_test_message(error_message)
            self.logger.error(f"이미지 생성 중 오류 발생: {str(e)}")
            return None

if __name__ == "__main__":
    today = datetime.now().strftime('%Y%m%d')
    
    # 로거 설정
    logger = LoggerUtil().get_logger()
    logger.info("==== 프로그램 시작 ====")
    
    if isTodayHoliday():
        logger.info('오늘은 공휴일입니다. 프로그램을 종료합니다.')
        sys.exit()

    # 전체 종목 정보 가져오기 
    logger.info("전체 종목 정보 가져오기 시작")
    kospi_tickers = stock.get_market_ticker_list(date=today, market="KOSPI")
    kosdaq_tickers = stock.get_market_ticker_list(date=today, market="KOSDAQ")

    telegram = TelegramUtil()
    api_util = ApiUtil()
    report = InstitutionTotalReport()
    
    # 기관 순매수 데이터 조회
    result = report.get_institution_total_report()
    
    # 상위 10개만 필터링
    filtered_data = result[:10] if len(result) > 10 else result

    logger.info("코스피 지수 조회 시작")
    # 코스피 지수 조회
    kospi_result = report.get_domestic_index(market_code="KOSPI", date=today)
    kospi_index_change_rate = round(((kospi_result.iloc[0]['종가'] - kospi_result.iloc[29]['종가']) / kospi_result.iloc[29]['종가'] * 100), 2)
    logger.info(f"코스피 지수 조회 완료: 30일간 등락률 {kospi_index_change_rate}%")

    logger.info("코스닥 지수 조회 시작")
    # 코스닥 지수 조회
    kosdaq_result = report.get_domestic_index(market_code="KOSDAQ", date=today)
    kosdaq_index_change_rate = round(((kosdaq_result.iloc[0]['종가'] - kosdaq_result.iloc[29]['종가']) / kosdaq_result.iloc[29]['종가'] * 100), 2)
    logger.info(f"코스닥 지수 조회 완료: 30일간 등락률 {kosdaq_index_change_rate}%")

    reference_date = kospi_result.iloc[29]['날짜'].strftime('%Y%m%d') # 한달 전 일자
    logger.info(f"과거 가격 조회 기준일: {reference_date}")
    
    # 기관 순매수 종목에 과거 가격 대비 등락률 정보 추가
    enhanced_data = report.add_historical_price_change(filtered_data, reference_date)
    
    # 시장 정보와 지수 등락률 추가
    final_data = report.add_market_info_and_index_rate(enhanced_data, kospi_index_change_rate, kosdaq_index_change_rate)

    df = report.convert_to_dataframe(final_data, top_n=10) # 상위 10개만 필터링하여 DataFrame으로 변환
    image_path = report.save_df_as_image(df) # DataFrame을 이미지로 저장
    
    if image_path:
        today = datetime.now().strftime('%Y-%m-%d')
        caption = f"{today} 기관 순매수 상위 TOP 10"
        
        logger.info("Telegram 메시지 전송 시작")
        telegram.send_multiple_photo([image_path], caption)
        logger.info("Telegram 메시지 전송 완료")
        
        try:
            logger.info("API 포스트 생성 시작")
            api_util.create_post(
                title=caption,
                content=f"{today} 기관 순매수 상위 TOP 10 결과",
                category="기관순매수",
                writer="admin",
                image_paths=[image_path],
                thumbnail_image_path=os.path.abspath("thumbnail/thumbnail.png")
            )
            logger.info("API 포스트 생성 완료")
        except ApiError as e:
            error_message = f"❌ API 오류 발생\n\n{e.message}"
            telegram.send_test_message(error_message)
            logger.error(f"API 포스트 생성 오류: {e.message}")
    else:
        logger.warning("이미지 생성에 실패했습니다.")
        
    logger.info("==== 프로그램 종료 ====")
    

