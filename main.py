import requests
from dotenv import load_dotenv
import os
import json
from datetime import datetime
import pandas as pd
import imgkit
from utils.api_util import ApiUtil, ApiError
from utils.telegram_util import TelegramUtil

load_dotenv()

class InstitutionTotalReport:
    def __init__(self):
        self.url_base = os.getenv("KIS_URL_BASE")
        self.app_key = os.getenv("KIS_APP_KEY")
        self.app_secret = os.getenv("KIS_APP_SECRET")
        self.token_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'token.json')
        self.img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
        self.wkhtmltoimage_path = os.getenv('WKHTMLTOIMAGE_PATH')
        
        # img 디렉토리가 없으면 생성
        if not os.path.exists(self.img_dir):
            os.makedirs(self.img_dir)

    def load_token(self):
        """토큰 파일에서 저장된 토큰 정보를 로드"""
        if not os.path.exists(self.token_file):
            return None
            
        with open(self.token_file, 'r') as f:
            data = json.load(f)
            
        # 만료 시간 확인 (둘 중 하나라도 만료되면 새로운 토큰 발급)
        now = datetime.now()
        expires_at = datetime.strptime(data['access_token_token_expired'], "%Y-%m-%d %H:%M:%S")
        
        if expires_at <= now:
            return None
            
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

    def check_env_variables(self):
        """필수 환경변수 체크"""
        required_vars = ['KIS_APP_KEY', 'KIS_APP_SECRET', 'KIS_URL_BASE']
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        
        if missing_vars:
            raise Exception(f"다음 환경변수가 설정되지 않았습니다: {', '.join(missing_vars)}")

    def get_token(self):
        """토큰 조회 또는 새로 발급"""
        # 환경변수 체크
        self.check_env_variables()
        
        # 저장된 토큰이 있는지 확인
        token = self.load_token()
        if token:
            return token

        # 새로운 토큰 발급
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
            raise Exception("토큰 발급 실패")
            
        token_info = res.json()
        self.save_token(token_info)
        
        return token_info['access_token'] 
    
    def get_institution_total_report(self):
        token = self.get_token()
        if not token:
            raise Exception("토큰 발급 실패")
            
        # API 엔드포인트 설정
        PATH = "uapi/domestic-stock/v1/quotations/foreign-institution-total"
        # PATH = "uapi/domestic-stock/v1/quotations/inquire-price"
        URL = f"{self.url_base}/{PATH}"

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
            return result
        else:
            raise Exception(f"API 호출 실패: {res.json()['msg_cd']}")
    
    def convert_to_dataframe(self, data, top_n=10):
        """API 응답 데이터를 DataFrame으로 변환"""
        if not data:
            return pd.DataFrame()
        
        # 상위 N개만 필터링
        filtered_data = data[:top_n] if len(data) > top_n else data
        
        # 필요한 컬럼 추출
        df = pd.DataFrame(filtered_data)
        
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
                
        result_df['전일대비율(%)'] = df['prdy_ctrt'].apply(format_rate)
        result_df['기관순매수량'] = df['orgn_ntby_qty'].astype(int).map('{:,}'.format)
        result_df['기관순매수금액'] = (df['orgn_ntby_tr_pbmn'].astype(float) / 100).round(2).map('{:,}'.format)  # 억원 단위로 변환
        
        return result_df
    
    def save_df_as_image(self, df, file_name="institution_top_report"):
        """DataFrame을 이미지로 저장하고 파일 경로 반환"""
        if df.empty:
            print("DataFrame이 비어 있어 이미지를 생성할 수 없습니다.")
            return None
            
        if not file_name.endswith('.png'):
            file_name = file_name + '.png'
            
        file_name, file_extension = os.path.splitext(file_name)
        current_date = datetime.now().strftime('%Y%m%d')
        new_file_path = os.path.join(self.img_dir, f"{file_name}_{current_date}{file_extension}")
        
        # 이전 파일 삭제
        for old_file in os.listdir(self.img_dir):
            if old_file.startswith(file_name) and old_file.endswith(file_extension):
                try:
                    os.remove(os.path.join(self.img_dir, old_file))
                    print(f"기존 파일 삭제: {old_file}")
                except:
                    pass

        # 캡션 설정
        today_display = datetime.now().strftime('%Y-%m-%d')
        caption = f"{today_display} 기관 순매수 상위 종목"

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
                raise ValueError("WKHTMLTOIMAGE_PATH 환경변수가 필요합니다.")
                
            config = imgkit.config(wkhtmltoimage=self.wkhtmltoimage_path)
            imgkit.from_string(html_str, new_file_path, options=options, config=config)
            print(f"새 파일 저장: {new_file_path}")
            
            return new_file_path
            
        except Exception as e:
            error_message = f"❌ 오류 발생\n\n함수: save_df_as_image\n파일: {file_name}\n오류: {str(e)}"
            print(f"이미지 생성 중 오류 발생: {str(e)}")
            return None
        
if __name__ == "__main__":
    telegram = TelegramUtil()
    api_util = ApiUtil()
    report = InstitutionTotalReport()
    result = report.get_institution_total_report()
    df = report.convert_to_dataframe(result, top_n=10) # 상위 10개만 필터링하여 DataFrame으로 변환
    image_path = report.save_df_as_image(df) # DataFrame을 이미지로 저장
    
    if image_path:
        today = datetime.now().strftime('%Y-%m-%d')
        caption = f"{today} 기관 순매수 상위 TOP 10"
        telegram.send_multiple_photo([image_path], caption)
        try:
            api_util.create_post(
                title=caption,
                content=f"{today} 기관 순매수 상위 TOP 10 결과",
                category="기관순매수",
                writer="admin",
                image_paths=[image_path]
            )
        except ApiError as e:
            error_message = f"❌ API 오류 발생\n\n{e.message}"
            telegram.send_test_message(error_message)
    else:
        print("이미지 생성에 실패했습니다.")
    

