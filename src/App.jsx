import React, { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
import _ from 'lodash';
import XLSX from 'xlsx-js-style';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  LineChart, Line, PieChart, Pie, Cell, AreaChart, Area, LabelList, ComposedChart
} from 'recharts';

import {
  LayoutDashboard, BarChart3, Package, Sparkles, ClipboardList, LogOut,
  Upload, TrendingUp, TrendingDown, Users, ChevronRight, FileText, RefreshCw, FileSpreadsheet, Calendar
} from 'lucide-react';

const COLORS = ['#666C59', '#8A9178', '#A3A895', '#10b981', '#f59e0b', '#ef4444'];
const THEME_COLOR = '#666C59'; // SkuView 메인 테마 컬러 (Olive Drab)
const THEME_BG = 'rgba(102, 108, 89, 0.08)'; // SkuView 활성화 배경색

// 상품번호(N열) 기준 '이쁜 이름 + 카테고리' 사전
const PRODUCT_DICTIONARY = {
  '13403180807': { name: '공방가죽 미니 삼각필통', category: '필통' },
  '13374663512': { name: '줄라이 미니 카드 지갑', category: '지갑' },
  '13346674237': { name: '커버형 클립보드 A4', category: '기타' },
  '13331055788': { name: '줄라이 아코디언 카드 지갑', category: '지갑' },
  '13312715270': { name: '공방클래식 가죽 골프지갑', category: '지갑' },
  '13142228329': { name: '줄라이 가죽 안경파우치', category: '안경파우치' },
  '13026176860': { name: '줄라이 여권케이스', category: '여권파우치' },
  '12967044389': { name: '공방가죽 성경책 커버', category: '북커버' },
  '12930661931': { name: '돛단배 천연소가죽 필통', category: '필통' },
  '12873783614': { name: '아테네 가죽 지퍼형 메모패드+속지SET', category: '메모패드' },
  '12867264028': { name: '줄라이 가죽 지퍼형 메모패드+속지SET', category: '메모패드' },
  '12769874967': { name: '천연소가죽 명함케이스', category: '파우치/케이스' },
  '12748043827': { name: '시그니처 양면 가죽 북마크', category: '북마크' },
  '12705670806': { name: '줄라이 가죽 메모패드+속지SET', category: '메모패드' },
  '12705440140': { name: '잉크펜 리필용심', category: '기타' },
  '12649661297': { name: '천연소가죽 북스토퍼', category: '북마크' },
  '12552768092': { name: '사피아노 여권케이스 슬리브형', category: '여권파우치' },
  '12491264921': { name: '사피아노 가죽 메모패드+속지SET', category: '메모패드' },
  '12421607787': { name: '줄라이 가죽 골프지갑', category: '지갑' },
  '12248300594': { name: '높이조절 2중관절 독서대 +가죽매트SET', category: '독서대' },
  '11887545683': { name: '헤링본 가죽 메모패드+속지SET', category: '메모패드' },
  '11755502213': { name: '[N배송] 휴대용 독서대+가죽파우치 세트', category: '독서대' },
  '11635279207': { name: '휴대용 독서대+공방가죽 삼각필통 SET', category: '독서대' },
  '11557665044': { name: '리갈패드 메모패드 속지', category: '메모패드' },
  '11557550116': { name: '잉크펜 리필가능', category: '기타' },
  '11554094489': { name: '공방클래식 가죽 메모패드+속지SET', category: '메모패드' },
  '11494442981': { name: '아테네 매일미사 책커버', category: '북커버' },
  '11291684196': { name: '아테네 가죽 북커버', category: '북커버' },
  '11097872946': { name: '공방가죽 삼각필통', category: '필통' },
  '11016420098': { name: '공방가죽 다용도 용돈봉투', category: '파우치/케이스' },
  '10994046250': { name: '휴대용 저소음슬림 독서대+가죽파우치SET', category: '독서대' },
  '10807630995': { name: '공방가죽 메모패드+속지SET', category: '메모패드' },
  '10769529101': { name: '아테네가죽 메모패드+속지SET', category: '메모패드' },
  '10760807810': { name: '공방 민음사 북커버', category: '북커버' },
  '10704047366': { name: '공방 클래식 북커버', category: '북커버' },
  '10685235249': { name: '선재업고튀어 대본집 북커버', category: '북커버' },
  '10306457578': { name: '토고 가죽 북커버', category: '북커버' },
  '9775830574': { name: '공방 가죽 북마크', category: '북마크' },
  '9581200417': { name: '휴대용 독서대+가죽파우치SET', category: '독서대' },
  '9580896783': { name: '티코스터', category: '기타' },
  '9457308268': { name: '사피아노 가죽 골프지갑', category: '지갑' },
  '9455231644': { name: '헤링본 가죽 골프지갑', category: '지갑' },
  '9455222718': { name: '프리미엄 가죽 골프지갑', category: '지갑' },
  '9435745083': { name: '공방 코끼리걸쇠 소품파우치', category: '파우치/케이스' },
  '9435728696': { name: '푸에블로 코끼리걸쇠 소품파우치', category: '파우치/케이스' },
  '9434640990': { name: '헤링본 가죽 다용도 장지갑', category: '지갑' },
  '9434295247': { name: '헤링본 미니 카드지갑', category: '지갑' },
  '9396550110': { name: '미니 가죽 북마크(3개1세트)', category: '북마크' },
  '9396453834': { name: '프리미엄 여권파우치', category: '여권파우치' },
  '9392368866': { name: '헤링본 여권 파우치', category: '여권파우치' },
  '9391960745': { name: '프리미엄 고리형 안경파우치', category: '안경파우치' },
  '9380333349': { name: '헤링본 안경파우치', category: '안경파우치' },
  '9102513826': { name: '사피아노 가죽 아코디언 지갑', category: '지갑' },
  '9102480003': { name: '헤링본 아코디언 지갑', category: '지갑' },
  '9096246199': { name: '헤링본 카드지갑(미니미니)', category: '지갑' },
  '9096008440': { name: '프리미엄 가죽 카드지갑(미니미니)', category: '지갑' },
  '8714740372': { name: '공방 가죽 북커버', category: '북커버' },
  '8701790242': { name: '헤링본 북커버', category: '북커버' },
  '8606275232': { name: '아코디언 카드지갑', category: '지갑' },
  '8315351864': { name: '사피아노 용돈봉투', category: '파우치/케이스' },
  '8308579934': { name: '공방 가죽 필통', category: '필통' },
  '8271595857': { name: '하프문 가죽 필통', category: '필통' },
  '8198969600': { name: '푸에블로 코끼리걸쇠 소품파우치', category: '파우치/케이스' },
  '8134352927': { name: '코끼리걸쇠 소품파우치', category: '파우치/케이스' },
  '6919881007': { name: '사피아노 북커버', category: '북커버' },
  '6918592921': { name: '사피아노 북커버', category: '북커버' },
  '6207665439': { name: '양면 밴딩 필통', category: '필통' },
  '6006763857': { name: '헤링본 데스크매트', category: '데스크매트' },
  '5962399798': { name: '양면 밴딩 필통', category: '필통' },
  '5931986099': { name: '프리미엄 북커버', category: '북커버' },
  '5857361440': { name: '헤링본 가죽 용돈봉투', category: '파우치/케이스' },
  '5854966728': { name: '다용도 멀티 소품파우치', category: '파우치/케이스' },
  '5738326723': { name: '마우스 사각패드', category: '기타' },
  '5723733361': { name: '슬림 필통 펜파우치', category: '필통' },
  '5701144232': { name: '프리미엄 안경파우치', category: '안경파우치' },
  '5686965533': { name: '다용도 지갑', category: '지갑' },
  '5683667686': { name: '프리미엄 미니 카드지갑', category: '지갑' },
  '5474342676': { name: '헤링본 북커버', category: '북커버' },
  '5303909572': { name: '노트북 가죽 파우치', category: '파우치/케이스' },
  '5303901477': { name: '워싱브러쉬 슬림 필통', category: '필통' },
  '5303896044': { name: '슬림 헤링본 필통', category: '필통' },
  '5303890952': { name: '클래식 슬림 필통', category: '필통' },
  '5303505977': { name: '조약돌 마우스패드', category: '기타' },
  '5301975478': { name: '프리미엄 가죽 데스크매트', category: '데스크매트' },
};

const simplifyName = (name, id) => {
  // 1순위: 사전에 등록된 이름이 있으면 즉시 반환
  if (id && PRODUCT_DICTIONARY[id]) {
    return PRODUCT_DICTIONARY[id].name;
  }

  if (!name) return 'Unknown';

  // 2순위: 사전에 없으면 규칙 기반 정제
  return name
    .replace(/\[.*?\]/g, '') // [각인가능] 등 대괄호 내용 삭제
    .replace(/오브리즈/g, '') // 브랜드명 삭제
    .replace(/맞춤형/g, '')  // 불필요한 수식어 삭제
    .replace(/리폼/g, '')    // 불필요한 키워드 삭제
    .replace(/\s+/g, ' ')   // 연속된 공백 하나로 축소
    .trim();
};

const getCategory = (id) => {
  if (id && PRODUCT_DICTIONARY[id]) return PRODUCT_DICTIONARY[id].category;
  return '기타';
};

const parseOptions = (optionStr) => {
  if (!optionStr || optionStr === '기본') return { color: '기본', size: '기본' };

  // 불필요한 띄어쓰기 및 괄호 내용 정제
  const cleanedStr = optionStr
    .replace(/\[.*?\]/g, '')            // 대괄호와 그 안의 내용 모두 삭제
    .replace(/\s*([()（）])\s*/g, '$1') // 괄호 앞뒤 공백 제거 (전각 괄호 포함)
    .replace(/\s*:\s*/g, ':')           // 콜론 앞뒤 공백 제거
    .replace(/\s*\/\s*/g, ' / ');        // 슬래시(구분자) 앞뒤는 깔끔하게 한 칸만

  const result = { color: '기본', size: '기본' };
  const parts = cleanedStr.split('/').map(p => p.trim());

  let tempChoice = '';
  let tempSize = '';
  let tempColors = [];

  parts.forEach(part => {
    if (part.includes(':')) {
      const [key, val] = part.split(':').map(s => s.trim());
      // 키의 공백을 제거하고 비교하여 인식률 향상
      if (key.replace(/\s+/g, '').match(/색상|컬러|Color|워싱펜케이스|헤링본펜케이스|미니카드지갑|헤링본데스크매트|노트북파우치|가죽용돈봉투|소품파우치/i)) {







        tempColors.push(val);
      } else if (key.match(/사이즈|규격|Size|세트/i)) {
        tempSize = val;

      } else if (key.match(/타입|Type|선택|상품|상픔/i)) {
        tempChoice = val;
      }
    } else {
      // 키워드가 명시되지 않은 경우 일부 색상 및 사이즈 패턴 매칭
      if (part.match(/블랙|화이트|네이비|레드|그린|블루|옐로우|브라운|그레이|카키|베이지/i)) {
        tempColors.push(part);
      } else if (part.match(/^[SML]$|^[0-9]{2,3}$|^[XSL]{2,3}$|옵션[0-9]{2}/i)) {
        tempSize = part;
      }
    }
  });

  // 색상이 여러 개일 경우 + 로 연결 (띄어쓰기 제거하여 통일)
  if (tempColors.length > 0) {
    result.color = tempColors.map(c => c.replace(/\s+/g, '')).join('+');
  }

  // 오픈형XL -> 오픈형(XL) 로 변환하는 헬퍼 함수
  const formatEmbeddedSize = (str) => {
    return str.replace(/^(.*?[가-힣0-9])\s*(XS|S|M|L|XL|XXL|2XL|3XL)$/i, (match, p1, p2) => {
      return `${p1}(${p2.toUpperCase()})`;
    });
  };


  tempChoice = formatEmbeddedSize(tempChoice);
  tempSize = formatEmbeddedSize(tempSize);

  if (tempChoice && tempSize) {
    result.size = `${tempChoice}(${tempSize})`;
  } else if (tempChoice) {
    result.size = tempChoice;
  } else if (tempSize) {
    result.size = tempSize;
  }

  return result;
};

function App() {
  const [allData, setAllData] = useState([]);
  const allDataRef = React.useRef([]);        // useEffect 이중 재계산 방지용 ref
  const [summary, setSummary] = useState(null);
  const [loading, setLoading] = useState(false);
  const [activeScreen, setActiveScreen] = useState('guide');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [totalDateRange, setTotalDateRange] = useState({ start: '', end: '' });
  const [drillDownInfo, setDrillDownInfo] = useState(null);


  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);

    // Naver Smart Store CSVs are typically EUC-KR, but some newer ones are UTF-8.
    // PapaParse handles most cases, but we'll stick to EUC-KR as primary for Korean Excel exports.
    Papa.parse(file, {
      encoding: 'EUC-KR',
      header: true,
      skipEmptyLines: true,
      transformHeader: (h) => h.trim(), // 헤더 명칭의 앞뒤 공백 제거
      complete: (results) => {
        try {
          if (!results.data || results.data.length === 0) {
            alert('파일에 데이터가 없습니다.');
            setLoading(false);
            return;
          }

          const cleaned = results.data
            .map(row => cleanData(row))
            .filter(row => row !== null && row.date !== null)
            .filter(d => {
              // 추가구성상품 중 제외 키워드 처리
              if (d.productType === '추가구성상품') {
                const excludeKeywords = ['내지', '선물포장', '북마커', '리필용심', 'GIFT', 'EVENT', '마우스패드', '티슈케이스', '골프지갑', '잉크펜'];
                const hasExcludeKeyword = excludeKeywords.some(k => d.option.includes(k));
                if (hasExcludeKeyword && !d.isEngraving) return false;
                if (!d.isEngraving) return false;
              }
              return true;
            });

          if (cleaned.length === 0) {
            const headers = Object.keys(results.data[0] || {}).join(', ');
            const sampleData = results.data.slice(0, 2).map(r => JSON.stringify(r)).join('\n');
            alert(`데이터를 분석할 수 없습니다.\n\n이 파일은 '상품 목록'이거나 필수 항목(결제일, 결제금액 등)이 누락된 것 같습니다.\n대시보드 분석을 위해서는 스마트스토어의 '판매 리포트' 또는 '주문 리포트' CSV 파일을 업로드해주세요.\n\n[감지된 헤더]: ${headers}\n\n[데이터 샘플]:\n${sampleData}`);
            setLoading(false);
            return;
          }

          allDataRef.current = cleaned;     // ref 먼저 동기 업데이트
          setAllData(cleaned);

          const allDates = cleaned.map(d => d.dateStr).filter(Boolean).sort();
          const startStr = allDates[0] || '';
          const endStr = allDates[allDates.length - 1] || '';

          setStartDate(startStr);
          setEndDate(endStr);
          setTotalDateRange({ start: startStr, end: endStr });

          // 메인 스레드 블로킹 방지: 다음 이벤트 루프에서 무거운 연산 실행
          setLoading(false);
          setTimeout(() => {
            setSummary(calculateSummary(cleaned));
            setActiveScreen('dashboard');
          }, 0);
        } catch (err) {
          console.error('Data processing error:', err);
          alert(`데이터 처리 중 오류가 발생했습니다: ${err.message}`);
          setLoading(false);
        }
      },
      error: () => {
        alert('파일을 읽는 중 오류가 발생했습니다.');
        setLoading(false);
      }
    });
  };

  // 날짜 변경 시 요약 데이터 재계산
  // allData 대신 allDataRef 사용 → 파일 업로드 직후 이중 재계산 방지
  useEffect(() => {
    const data = allDataRef.current;
    if (data.length > 0 && startDate && endDate) {
      const filtered = data.filter(d => d.dateStr >= startDate && d.dateStr <= endDate);
      setSummary(calculateSummary(filtered));
    }
  }, [startDate, endDate]);

  const cleanData = (row) => {
    try {
      if (!row || typeof row !== 'object') return null;
      // 헤더 유연성 확보 (스마트스토어의 다양한 리포트 형식 대응)
      const productType = row['상품종류'] || row['구분'] || row['종류'] || '';
      const optionInfo = row['옵션정보'] || row['옵션'] || '';
      const dateStr = row['구매확정일'] || row['결제일'] || row['주문일시'] || row['판매일시'] || row['상품등록일'] || '';
      const qtyStr = row['수량'] || row['판매수량'] || '1';
      const revenueStr = row['정산예정금액'] || row['결제금액'] || row['판매금액'] || row['판매가'] || '0';
      const productName = row['상품명'] || '';
      const productNumber = row['상품번호'] || row['상품번호(스마트스토어)'] || '';
      const device = row['결제위치'] || row['유입경로'] || 'Unknown';
      const buyerId = row['구매자ID'] || row['구매자명'] || 'Unknown';
      const orderId = row['주문번호'] || row['상품주문번호'] || row['주문번호'] || '';
      const optionPriceStr = row['옵션가격'] || '0';

      const simplifiedProduct = simplifyName(productName, productNumber);
      let { color, size } = parseOptions(optionInfo);

      // 공방 가죽 북커버(8714740372) 전용 사이즈 옵션 매칭 (이력 통합)
      const targetProdNum = String(productNumber).trim();
      
      if (targetProdNum === '8714740372') {
        const s = size.replace(/\s+/g, ''); // 공백 제거 후 비교 (정확도 향상)
        if (s === '46판B6(일반형)' || s === '46판B6（일반형）') size = '일반형(S)46판B6';
        else if (s === '국판A5(일반형)' || s === '국판A5（일반형）') size = '일반형(M)국판A5';
        else if (s === '46배판(일반형)' || s === '46배판（일반형）' || s === '일반형(L)46배판') size = '일반형(L)신국판';
        else if (s === '민음사(일반형)' || s === '민음사（일반형）') size = '일반형(민음사)';
        else if (s === '신46판(오픈형)' || s === '신46판（오픈형）') size = '오픈형(S)신46판';
        else if (s === '신국판(오픈형)' || s === '신국판（오픈형）') size = '오픈형(L)신국판';
        else if (s === '민음사(오픈형)' || s === '민음사（오픈형）') size = '오픈형(민음사)';
      }

      // 특정 제품군: 사이즈 옵션 공백 제거 및 전각 문자 치환 (통합 관리)
      const needsNormalization = ['11291684196', '10306457578', '10760807810'];
      if (needsNormalization.includes(targetProdNum)) {

        size = size.replace(/\s+/g, '')
                   .replace(/（/g, '(')
                   .replace(/）/g, ')')
                   .replace(/Ｓ/g, 'S')
                   .replace(/Ｍ/g, 'M')
                   .replace(/Ｌ/g, 'L')
                   .replace(/Ｘ/g, 'X');
      }

      // 토고 가죽 북커버(10306457578) 전용 추가 매칭
      if (targetProdNum === '10306457578') {
        if (size === '46배판(일반형)') size = '신국판(일반형)';
      }







      // 숫자 데이터 정제
      const revenue = typeof revenueStr === 'string'
        ? parseInt(revenueStr.replace(/[^0-9-]/g, '')) || 0
        : parseInt(revenueStr) || 0;
      const qty = typeof qtyStr === 'string'
        ? parseInt(qtyStr.replace(/[^0-9]/g, '')) || 0
        : parseInt(qtyStr) || 0;
      const optionPrice = typeof optionPriceStr === 'string'
        ? parseInt(optionPriceStr.replace(/[^0-9-]/g, '')) || 0
        : parseInt(optionPriceStr) || 0;

      // 각인 판별 로직 고도화
      const hasEngravingInOption = optionInfo.includes('각인신청:') && !optionInfo.includes('각인신청안함');
      const isSeparateEngravingItem = productType === '추가구성상품' && optionInfo.includes('각인');
      const isEngraving = isSeparateEngravingItem || hasEngravingInOption;

      // 각인 매출액 계산 (조합형은 옵션가격을, 추가구성은 전체금액을 사용)
      let engravingAmount = 0;
      if (isSeparateEngravingItem) {
        engravingAmount = revenue;
      } else if (hasEngravingInOption) {
        engravingAmount = optionPrice * qty;
      }

      // 날짜 파싱 및 형식화 (로컬 기준)
      let date = null;
      let month = 'Unknown';
      let dStr = '';

      if (dateStr) {
        // 1. 문자열 변환 및 양끝 공백 제거
        const raw = dateStr.toString().trim();
        // 2. 점(.)을 대시(-)로 변경
        let normalized = raw.replace(/\./g, '-');
        // 3. 날짜와 시간 사이의 공백이나 'T'를 기준으로 첫 번째 파트(날짜)만 추출
        const datePart = normalized.split(/[\sT]/)[0];
        // 4. 숫자와 대시만 남김 (예: "2026-04-01")
        let finalDPart = datePart.replace(/[^0-9-]/g, '');
        if (finalDPart.endsWith('-')) finalDPart = finalDPart.slice(0, -1);

        // YYYY-MM-DD 형식으로 보정
        const parts = finalDPart.split('-');
        if (parts.length === 3) {
          const y = parts[0];
          const m = parts[1].padStart(2, '0');
          const d = parts[2].padStart(2, '0');
          dStr = `${y}-${m}-${d}`;
          date = new Date(dStr);

          if (!isNaN(date.getTime())) {
            month = `${y}-${m}`;
          } else {
            date = null;
            dStr = '';
          }
        }
      }

      return {
        date,
        month,
        dateStr: dStr,
        revenue, // 전체 결제 금액 (검증용)
        productRevenue: revenue, // 각인비 포함 전체 매출액 사용 (사용자 요청)
        engravingAmount: engravingAmount,
        product: isSeparateEngravingItem ? '각인 서비스' : simplifiedProduct,
        category: isSeparateEngravingItem ? '각인' : getCategory(productNumber),
        productNumber,
        option: optionInfo || '기본',
        color,
        size,
        device,
        buyerId,
        orderId,
        productType,
        isEngraving,
        sku: isSeparateEngravingItem ? '각인 서비스' : `${simplifiedProduct} ${optionInfo}`.trim()
      };

    } catch (err) {
      return null;
    }
  };

  const calculateSummary = (cleanedData) => {
    // 본 상품과 각인 매출을 모두 포함하여 전체 매출 산출
    const mainSales = cleanedData.filter(d => d.productType === '조합형옵션상품' || d.isEngraving);
    const engravingSales = cleanedData.filter(d => d.isEngraving);

    const totalRevenue = _.sumBy(mainSales, 'revenue');

    const totalQty = _.sumBy(mainSales, 'qty');
    const engravingTotalRevenue = _.sumBy(cleanedData, 'engravingAmount');
    const engravingTotalQty = _.sumBy(engravingSales, 'qty');
    
    // 각인 채택률 계산 (전체 주문번호 중 각인 주문번호 비중)
    const totalOrderIds = new Set(mainSales.map(d => d.orderId));
    const engravingOrderIds = new Set(engravingSales.map(d => d.orderId));
    const engravingRate = totalOrderIds.size > 0 ? (engravingOrderIds.size / totalOrderIds.size) * 100 : 0;

    const engravingTrend = _(engravingSales)
      .groupBy('month')
      .map((items, month) => ({ month, revenue: _.sumBy(items, 'revenue'), qty: _.sumBy(items, 'qty') }))
      .sortBy('month').value();


    const deviceStats = _(mainSales)
      .groupBy('device')
      .map((items, device) => ({ name: device, value: _.sumBy(items, 'productRevenue') }))
      .value();

    // 고객별 첫 주문 ID 식별
    const ordersByBuyer = _(mainSales)
      .groupBy('buyerId')
      .mapValues(items => _.sortBy(_.uniqBy(items, 'orderId'), 'date'))
      .value();

    const firstOrderIds = new Set();
    Object.values(ordersByBuyer).forEach(orders => {
      if (orders.length > 0) {
        firstOrderIds.add(orders[0].orderId);
      }
    });

    // 각인 월별 집계 맵 미리 생성
    const engravingByMonth = _(engravingSales)
      .groupBy('month')
      .mapValues(items => _.sumBy(items, 'engravingAmount'))
      .value();

    const monthlyTrend = _(mainSales)
      .groupBy('month')
      .map((items, month) => {
        const prodRev = _.sumBy(items, 'productRevenue');
        const repurchaseRev = _.sumBy(items, item => firstOrderIds.has(item.orderId) ? 0 : item.productRevenue);
        return { 
          month, 
          revenue: prodRev, 
          repurchaseRevenue: repurchaseRev,
          repurchaseShare: prodRev > 0 ? (repurchaseRev / prodRev) * 100 : 0,
          engravingRevenue: engravingByMonth[month] || 0,
          qty: _.sumBy(items, 'qty') 
        };
      })
      .sortBy('month').value();



    // 1. 제품별 매출액 집계 (각인비 포함)
    const rawProductStats = _(mainSales)
      .filter(d => d.product !== '각인 서비스')
      .groupBy('product')
      .map((items, product) => {
        return { 
          name: product, 
          revenue: _.sumBy(items, 'revenue'), 
          qty: _.sumBy(items, 'qty'),
          category: items[0]?.category || '기타'
        };
      })
      .value();

    // 2. 별도 품목으로 결제된 각인 서비스 항목 추가
    const separateEngraving = mainSales.filter(d => d.product === '각인 서비스');
    if (separateEngraving.length > 0) {
      rawProductStats.push({
        name: '각인 서비스(개별결제)',
        revenue: _.sumBy(separateEngraving, 'revenue'),
        qty: _.sumBy(separateEngraving, 'qty'),
        category: '각인'
      });
    }

    // 3. 정렬 및 비중/누적비중 계산
    let cumulative = 0;
    const productStats = _(rawProductStats)
      .sortBy('revenue')
      .reverse()
      .map(p => {
        const share = totalRevenue > 0 ? (p.revenue / totalRevenue) * 100 : 0;
        const prevCumulative = cumulative;
        cumulative += share;
        return { 
          ...p, 
          share,
          prevCumulative,
          cumulativeShare: Math.min(100, cumulative) // 부동소수점 오차 방지
        };
      })
      .value();


    const skuStats = _(mainSales).groupBy('sku').map((items, sku) => ({ name: sku, revenue: _.sumBy(items, 'productRevenue'), qty: _.sumBy(items, 'qty') })).sortBy('revenue').reverse().take(20).value();
    
    // 카테고리별 집계 (일관성을 위해 d.category || '기타' 사용)
    const categoryStats = _(mainSales).groupBy(d => d.category || '기타').map((items, category) => ({ 
      name: category, 
      revenue: _.sumBy(items, 'productRevenue'), 
      qty: _.sumBy(items, 'qty') 
    })).sortBy('revenue').reverse().value();

    const categoryProductMap = _(mainSales).groupBy(d => d.category || '기타').mapValues(items => 
      _(items).groupBy('product').map((pItems, name) => ({ 
        name: name || '알 수 없는 상품', 
        revenue: _.sumBy(pItems, 'productRevenue'), 
        qty: _.sumBy(pItems, 'qty'),
        category: pItems[0]?.category || '기타'
      })).sortBy('revenue').reverse().value()
    ).value();

    const colorStats = _(mainSales).groupBy(d => d.color || '기본').map((items, color) => ({ name: color, revenue: _.sumBy(items, 'productRevenue'), qty: _.sumBy(items, 'qty') })).sortBy('revenue').reverse().value();
    const sizeStats = _(mainSales).groupBy(d => d.size || '기본').map((items, size) => ({ name: size, revenue: _.sumBy(items, 'productRevenue'), qty: _.sumBy(items, 'qty') })).sortBy('revenue').reverse().value();

    // 고객별 구매 데이터 (고유 주문 기준)
    const customerPurchases = _(mainSales)
      .groupBy('buyerId')
      .map((items, id) => {
        const uniqueOrders = _.uniqBy(items, 'orderId');
        return { id, count: uniqueOrders.length, totalRevenue: _.sumBy(items, 'productRevenue') };
      })
      .value();

    const totalCustomers = customerPurchases.length;
    const totalOrders = _.uniqBy(mainSales, 'orderId').length;

    // 신규 주문: 각 고객의 첫 번째 주문 (데이터셋 내 기준)
    const newOrders = totalCustomers;
    // 재구매 주문: 전체 주문 중 첫 주문을 제외한 나머지 주문들
    const repeatOrders = Math.max(0, totalOrders - newOrders);
    const repurchaseRate = totalOrders > 0 ? (repeatOrders / totalOrders) * 100 : 0;

    const allDates = cleanedData.map(d => d.date).filter(Boolean);
    const minDate = allDates.length ? new Date(Math.min(...allDates)) : null;
    const maxDate = allDates.length ? new Date(Math.max(...allDates)) : null;
    const formatDate = (d) => d ? `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}` : '-';

    return {
      totalRevenue, totalQty, engravingTotalRevenue, engravingTotalQty, engravingRate, engravingTrend, monthlyTrend, productStats, skuStats, colorStats, sizeStats, deviceStats,

      repurchaseStats: { 
        totalCustomers, 
        repurchaseRate, 
        totalOrders, 
        newOrders, 
        repeatOrders, 
        loyalCustomers: _.sortBy(customerPurchases, 'count').reverse().slice(0, 10) 
      },
      dateRange: { start: formatDate(minDate), end: formatDate(maxDate) },
      activeProductsCount: _.uniqBy(mainSales, 'product').length,
      categoryStats, categoryProductMap,
      productOptionMap: _(mainSales).groupBy('product').mapValues(items => ({
        sizeStats: _(items).groupBy('size').map((s, name) => ({ name, qty: _.sumBy(s, 'qty'), revenue: _.sumBy(s, 'productRevenue') })).filter(s => s.name && s.name !== '기본').sortBy('qty').reverse().value(),
        colorStats: _(items).groupBy('color').map((c, name) => ({ name, qty: _.sumBy(c, 'qty'), revenue: _.sumBy(c, 'productRevenue') })).filter(c => c.name && c.name !== '기본').sortBy('qty').reverse().value(),
        // 사이즈와 색상을 조합한 상세 통계 추가
        combinedStats: _(items).groupBy(d => `${d.size}|${d.color}`).map((group, key) => {
          const [size, color] = key.split('|');
          return { size, color, qty: _.sumBy(group, 'qty'), revenue: _.sumBy(group, 'productRevenue') };
        }).sortBy('revenue').reverse().value(),
        monthlyTrend: _(items).groupBy('month').map((m, month) => ({ month, revenue: _.sumBy(m, 'productRevenue'), qty: _.sumBy(m, 'qty') })).sortBy('month').value()
      })).value()
    };
  };

  return (
    <div className="layout-root">
      {/* Sidebar */}
      <aside className="sidebar">
        <div className="sidebar-brand">
          <h2>판매 분석</h2>
          <p style={{ fontSize: '10px', color: 'rgba(255,255,255,0.5)' }}>스마트스토어 데이터 대시보드</p>
        </div>

        <nav className="sidebar-nav">
          <NavItem icon={<LayoutDashboard size={20} />} label="대시보드" active={activeScreen === 'dashboard'} onClick={() => setActiveScreen('dashboard')} />
          <NavItem icon={<BarChart3 size={20} />} label="카테고리 분석" active={activeScreen === 'product'} onClick={() => setActiveScreen('product')} />
          <NavItem icon={<ClipboardList size={20} />} label="옵션별 현황" active={activeScreen === 'sku'} onClick={() => { setDrillDownInfo(null); setActiveScreen('sku'); }} />
          <NavItem icon={<Sparkles size={20} />} label="가이드" active={activeScreen === 'guide'} onClick={() => setActiveScreen('guide')} />
        </nav>

        <div className="sidebar-footer" style={{ marginTop: 'auto' }}>
          <div className="nav-item">
            <LogOut size={20} />
            <span>로그아웃</span>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="main-content">
        {activeScreen === 'dashboard' && summary && (
          <DashboardView
            summary={summary}
            onUpload={handleFileUpload}
            startDate={startDate}
            endDate={endDate}
            onApplyFilter={(s, e) => {
              setStartDate(s);
              setEndDate(e);
            }}
            totalDateRange={totalDateRange}
            onDrillDown={(info) => {
              setDrillDownInfo(info);
              setActiveScreen('sku');
            }}
          />
        )}
        {activeScreen === 'product' && summary && (
          <ProductView 
            summary={summary} 
            startDate={startDate} 
            endDate={endDate} 
          />
        )}
        {activeScreen === 'sku' && summary && (
          <SkuView 
            summary={summary} 
            drillDownInfo={drillDownInfo} 
            onClearDrillDown={() => setDrillDownInfo(null)}
            startDate={startDate}
            endDate={endDate}
            onBack={() => setActiveScreen('product')}
          />
        )}
        {activeScreen === 'guide' && (
          <GuideView 
            summary={summary} 
            onUpload={handleFileUpload}
            loading={loading}
          />
        )}
        
        {!summary && activeScreen !== 'guide' && (
          <div style={{ 
            height: '80vh', 
            display: 'flex', 
            flexDirection: 'column', 
            alignItems: 'center', 
            justifyContent: 'center',
            textAlign: 'center'
          }}>
            <div style={{ 
              width: '80px', 
              height: '80px', 
              backgroundColor: '#eff6ff', 
              borderRadius: '50%', 
              display: 'flex', 
              alignItems: 'center', 
              justifyContent: 'center',
              marginBottom: '24px'
            }}>
              <Upload size={40} color="#3b82f6" />
            </div>
            <h2 style={{ fontSize: '24px', fontWeight: '800', marginBottom: '12px' }}>분석할 데이터가 없습니다</h2>
            <p style={{ color: '#666', marginBottom: '32px' }}>먼저 가이드 페이지에서 판매 데이터(CSV)를 업로드해주세요.</p>
            <button 
              onClick={() => setActiveScreen('guide')}
              className="btn-primary"
              style={{ padding: '12px 32px' }}
            >
              가이드 및 업로드로 이동
            </button>
          </div>
        )}
      </main>
    </div>
  );
}



// Sub-components
function NavItem({ icon, label, active, onClick }) {
  return (
    <div className={`nav-item ${active ? 'active' : ''}`} onClick={onClick}>
      {icon}
      <span>{label}</span>
    </div>
  );
}

function UploadScreen({ onUpload, loading }) {
  return (
    <div className="flex flex-col items-center justify-center h-[80vh]">
      <div className="card w-full max-w-2xl text-center py-20">
        <div className="w-20 h-20 bg-primary/10 rounded-full flex items-center justify-center mx-auto mb-6">
          <Upload size={40} className="text-primary" />
        </div>
        <h2 className="text-2xl font-bold mb-2">데이터 업로드 및 대시보드</h2>
        <p className="text-on-surface-variant mb-8">분석할 스마트스토어 판매 데이터 파일을 선택하세요.</p>
        <label className="btn-primary mx-auto w-fit">
          파일 선택하기
          <input type="file" hidden accept=".csv" onChange={onUpload} />
        </label>
        <p className="mt-6 text-xs text-on-surface-variant">최대 50MB 용량의 데이터 파일을 분석할 수 있습니다.</p>
      </div>
      {loading && <p className="mt-4 font-bold text-primary animate-pulse">다니엘이 분석 중입니다...</p>}
    </div>
  );
}

function DashboardView({ summary, onUpload, startDate, endDate, onApplyFilter, totalDateRange, onDrillDown }) {
  const handleExportExcel = () => {
    if (!summary || !summary.productStats) return;

    const wb = XLSX.utils.book_new();

    // 공통 스타일 설정 함수
    const applyCommonStyles = (worksheet, headerRowIndex, columns, isSummary = false) => {
      const getExcelColor = (prevCum) => {
        if (prevCum < 70) return '0078FF'; // A: Blue (70%를 넘기게 만드는 상품까지)
        if (prevCum < 80) return 'F97316'; // B: Orange
        if (prevCum < 90) return '22C55E'; // C: Green
        return 'EF4444'; // D: Red
      };

      const range = XLSX.utils.decode_range(worksheet['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[addr]) continue;

          let s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { vertical: 'center', horizontal: 'center' } };

          if (R === 0) {
            s = { font: { name: '맑은 고딕', sz: 16, bold: true, underline: true }, alignment: { horizontal: 'left' } };
          } else if (R >= 1 && R < headerRowIndex - 1) {
            s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'left' } };
          } else if (R === headerRowIndex) {
            s = {
              font: { name: '맑은 고딕', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
              fill: { fgColor: { rgb: '333333' } },
              alignment: { horizontal: 'center' },
              border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
            };
          } else if (R > headerRowIndex) {
            s.border = {
              top: { style: 'thin', color: { rgb: 'CCCCCC' } },
              bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
              left: { style: 'thin', color: { rgb: 'CCCCCC' } },
              right: { style: 'thin', color: { rgb: 'CCCCCC' } }
            };
            
            // 숫자 및 우측 정렬 처리
            if (worksheet[addr].t === 'n') {
              s.alignment = { horizontal: 'right' };
              worksheet[addr].z = '#,##0';
            }

            // [수정] 전체요약 시트의 누적비중 컬럼(C=3) 색상 적용 (이전 누적치 기준)
            if (isSummary && C === 3) {
              const itemIndex = R - (headerRowIndex + 1);
              if (summary.productStats[itemIndex]) {
                const prevCum = summary.productStats[itemIndex].prevCumulative;
                s.fill = { fgColor: { rgb: getExcelColor(prevCum) } };
                s.font = { ...s.font, color: { rgb: 'FFFFFF' }, bold: true };
              }
            }
          }
          worksheet[addr].s = s;
        }
      }
      worksheet['!cols'] = columns;
    };

    // 1. 메인 요약 시트 (기존)
    const title = [['주요 제품별 성과 분석 보고서']];
    const period = [[`분석 기간: ${startDate.replace(/-/g, '.')} ~ ${endDate.replace(/-/g, '.')}`]];
    const totalRev = [[`총 매출액: ${summary.totalRevenue.toLocaleString()}원`]];
    const summaryHeaders = [['제품명', '매출액(원)', '비중(%)', '누적비중(%)']];
    const summaryDataRows = summary.productStats.map(item => [
      item.name,
      item.revenue,
      `${item.share.toFixed(1)}%`,
      `${item.cumulativeShare.toFixed(1)}%`,
    ]);

    const wsSummary = XLSX.utils.aoa_to_sheet([...title, ...period, ...totalRev, [[]], ...summaryHeaders, ...summaryDataRows]);
    applyCommonStyles(wsSummary, 4, [{ wch: 45 }, { wch: 18 }, { wch: 10 }, { wch: 12 }], true);
    XLSX.utils.book_append_sheet(wb, wsSummary, '전체요약');

    // 2. 누적 비중 구간별 상세 시트 (A, B, C, D)
    const groups = [
      { id: 'A', name: 'A그룹(핵심_~70%)', min: 0, max: 70 },
      { id: 'B', name: 'B그룹(중점_~80%)', min: 70, max: 80 },
      { id: 'C', name: 'C그룹(일반_~90%)', min: 80, max: 90 },
      { id: 'D', name: 'D그룹(비주력_~100%)', min: 90, max: 101 }
    ];

    groups.forEach(group => {
      // prevCumulative를 기준으로 구간을 나누어 중복을 방지하고 임계값 초과 상품을 포함시킴
      const groupProducts = summary.productStats.filter(p => p.prevCumulative >= group.min && p.prevCumulative < group.max);
      if (groupProducts.length === 0) return;

      const groupTotalRev = _.sumBy(groupProducts, 'revenue');
      const groupTitle = [[`[${group.name}] 상세 성과 분석`]];
      const groupInfo = [[`분석 기간: ${startDate.replace(/-/g, '.')} ~ ${endDate.replace(/-/g, '.')}`]];
      const groupRevStr = [[`그룹 합계 매출: ${groupTotalRev.toLocaleString()}원 (전체의 ${(groupTotalRev / summary.totalRevenue * 100).toFixed(1)}%)`]];
      const groupHeaders = [['제품명', '사이즈', '색상', '매출액(원)', '비중(%)']];

      const groupRows = [];
      groupProducts.forEach(p => {
        const options = summary.productOptionMap[p.name]?.combinedStats || [{ size: '기본', color: '기본', revenue: p.revenue }];
        options.forEach(opt => {
          groupRows.push([
            p.name,
            opt.size,
            opt.color,
            opt.revenue,
            `${summary.totalRevenue > 0 ? (opt.revenue / summary.totalRevenue * 100).toFixed(1) : 0}%`
          ]);
        });
      });

      const wsGroup = XLSX.utils.aoa_to_sheet([...groupTitle, ...groupInfo, ...groupRevStr, [[]], ...groupHeaders, ...groupRows]);
      applyCommonStyles(wsGroup, 4, [{ wch: 35 }, { wch: 20 }, { wch: 20 }, { wch: 15 }, { wch: 10 }]);
      XLSX.utils.book_append_sheet(wb, wsGroup, `${group.id}그룹`);
    });

    const today = new Date();
    const d = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `판매분석_통합보고서_${d}.xlsx`);
  };


  const [tempStart, setTempStart] = useState(startDate);
  const [tempEnd, setTempEnd] = useState(endDate);

  // 파일 업로드 등으로 상위 startDate가 변경되면 연동
  useEffect(() => {
    setTempStart(startDate);
    setTempEnd(endDate);
  }, [startDate, endDate]);

  return (
    <div className="animate-fade-in">
      <div className="flex justify-between items-end mb-10">
        <div className="flex items-center gap-8">
          <h1 className="text-3xl font-bold">데이터 대시보드</h1>
          {totalDateRange.start && (
            <span className="text-xs font-normal text-on-surface-variant bg-surface-container px-3 py-1 rounded-full flex items-center gap-2">
              <FileText size={12} />
              파일 전체 기간: {totalDateRange.start.replace(/-/g, '.')} ~ {totalDateRange.end.replace(/-/g, '.')}
            </span>
          )}
        </div>
        <label className="btn-primary">
          <Upload size={18} />
          데이터 갱신
          <input type="file" hidden accept=".csv" onChange={onUpload} />
        </label>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '24px', marginBottom: '40px' }}>
        <div className="card stat-card" style={{ gridColumn: 'span 1' }}>
          <div className="flex flex-col h-full justify-between">
            <div>
              <p className="text-lg font-bold mb-8">분석 기간 설정</p>
              <div className="flex items-center gap-4 mb-4">
                <div className="flex-1 flex flex-col gap-2">
                  <label className="text-sm text-on-surface-variant font-bold">시작일</label>
                  <input
                    type="date"
                    value={tempStart}
                    onChange={(e) => setTempStart(e.target.value)}
                    min={totalDateRange.start}
                    max={totalDateRange.end}
                    style={{
                      width: '100%', padding: '12px 14px', borderRadius: '8px', border: '1px solid var(--outline-variant)',
                      fontSize: '15px', outline: 'none', backgroundColor: 'var(--surface-container-low)',
                      color: 'var(--on-surface)'
                    }}
                  />
                </div>
                <div className="mt-8 text-on-surface-variant opacity-40">~</div>
                <div className="flex-1 flex flex-col gap-2">
                  <label className="text-sm text-on-surface-variant font-bold">종료일</label>
                  <input
                    type="date"
                    value={tempEnd}
                    onChange={(e) => setTempEnd(e.target.value)}
                    min={totalDateRange.start}
                    max={totalDateRange.end}
                    style={{
                      width: '100%', padding: '12px 14px', borderRadius: '8px', border: '1px solid var(--outline-variant)',
                      fontSize: '15px', outline: 'none', backgroundColor: 'var(--surface-container-low)',
                      color: 'var(--on-surface)'
                    }}
                  />
                </div>
              </div>
            </div>
            <button
              className="btn-primary w-full py-2 text-sm justify-center"
              onClick={() => {
                if (tempStart > tempEnd) {
                  alert('시작일이 종료일보다 늦을 수 없습니다.');
                  return;
                }
                onApplyFilter(tempStart, tempEnd);
              }}
            >
              확인
            </button>
          </div>
        </div>
        <SummaryCard title="총 매출액 (전체)" value={`₩${summary.totalRevenue.toLocaleString()}`} trend="구매확정 기준" up titleClassName="text-lg font-bold mb-3" />
        <div className="card stat-card" style={{ display: 'flex', flexDirection: 'column' }}>
          <p className="text-lg font-bold mb-5">판매 데이터 분석</p>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', flex: 1 }}>
            <div style={{ backgroundColor: 'var(--surface-container-low)', padding: '12px', borderRadius: '12px', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
              <div className="flex items-center gap-1.5 mb-1 text-on-surface-variant opacity-60">
                <Package size={14} />
                <span className="text-[10px] font-bold uppercase tracking-wider">총 건수</span>
              </div>
              <span className="text-xl font-black">{summary.repurchaseStats.totalOrders.toLocaleString()}<small className="text-xs ml-0.5 opacity-50">건</small></span>
            </div>

            <div style={{ backgroundColor: 'var(--surface-container-low)', padding: '12px', borderRadius: '12px', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
              <div className="flex items-center gap-1.5 mb-1 text-on-surface-variant opacity-60">
                <Users size={14} />
                <span className="text-[10px] font-bold uppercase tracking-wider">신규주문</span>
              </div>
              <span className="text-xl font-black">{summary.repurchaseStats.newOrders.toLocaleString()}<small className="text-xs ml-0.5 opacity-50">건</small></span>
            </div>

            <div style={{ backgroundColor: 'rgba(var(--primary-rgb), 0.05)', padding: '12px', borderRadius: '12px', display: 'flex', flexDirection: 'column', justifyContent: 'center', border: '1px solid rgba(var(--primary-rgb), 0.1)' }}>
              <div className="flex items-center gap-1.5 mb-1 text-primary opacity-80">
                <RefreshCw size={14} />
                <span className="text-[10px] font-bold uppercase tracking-wider">재구매</span>
              </div>
              <span className="text-xl font-black text-primary">{summary.repurchaseStats.repeatOrders.toLocaleString()}<small className="text-xs ml-0.5 opacity-50">건</small></span>
            </div>

            <div style={{ backgroundColor: summary.repurchaseStats.repurchaseRate > 25 ? 'rgba(16, 185, 129, 0.05)' : 'rgba(239, 68, 68, 0.05)', padding: '12px', borderRadius: '12px', display: 'flex', flexDirection: 'column', justifyContent: 'center', border: summary.repurchaseStats.repurchaseRate > 25 ? '1px solid rgba(16, 185, 129, 0.1)' : '1px solid rgba(239, 68, 68, 0.1)' }}>
              <div className="flex items-center gap-1.5 mb-1" style={{ color: summary.repurchaseStats.repurchaseRate > 25 ? '#10b981' : '#ef4444', opacity: 0.8 }}>
                <TrendingUp size={14} />
                <span className="text-[10px] font-bold uppercase tracking-wider">재구매율</span>
              </div>
              <span className="text-xl font-black" style={{ color: summary.repurchaseStats.repurchaseRate > 25 ? '#10b981' : '#ef4444' }}>
                {summary.repurchaseStats.repurchaseRate.toFixed(1)}<small className="text-xs ml-0.5 opacity-50">%</small>
              </span>
            </div>
          </div>
        </div>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
        <div className="card">
          <h3 className="text-lg font-bold mb-6">월별 매출 트렌드</h3>
          <div style={{ height: '420px' }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={summary.monthlyTrend} margin={{ top: 30, right: 20, left: 30, bottom: 40 }}>

                <defs>
                  <linearGradient id="colorRev" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="var(--primary)" stopOpacity={0.12} />
                    <stop offset="95%" stopColor="var(--primary)" stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="colorRepurchase" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#fb923c" stopOpacity={0.12} />
                    <stop offset="95%" stopColor="#fb923c" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="month" axisLine={false} tickLine={false} fontSize={11} tickMargin={25} />
                <YAxis axisLine={false} tickLine={false} fontSize={11} tickMargin={15} tickFormatter={val => `${(val / 10000).toFixed(0)}만`} />
                <Tooltip formatter={(value) => [`₩${value.toLocaleString()}`, '']} />
                <Legend verticalAlign="top" align="right" height={36} iconType="circle" />
                
                {/* 전체 매출 */}
                <Area
                  name="전체 매출"
                  type="monotone"
                  dataKey="revenue"
                  stroke="var(--primary)"
                  fillOpacity={1}
                  fill="url(#colorRev)"
                  strokeWidth={2.5}
                  dot={{ r: 4, fill: 'var(--primary)', strokeWidth: 0 }}
                  activeDot={{ r: 6 }}
                >
                  <LabelList
                    dataKey="revenue"
                    position="top"
                    fontSize={12}
                    fontWeight={700}
                    formatter={(v) => `${(v / 10000).toFixed(0)}만`}
                    style={{ fill: '#0050cb' }}
                  />
                </Area>

                {/* 재구매 매출 */}
                <Area
                  name="재구매 매출"
                  type="monotone"
                  dataKey="repurchaseRevenue"
                  stroke="#fb923c"
                  fillOpacity={1}
                  fill="url(#colorRepurchase)"
                  strokeWidth={2}
                  dot={{ r: 3, fill: '#fb923c', strokeWidth: 0 }}
                >
                  <LabelList
                    dataKey="repurchaseRevenue"
                    content={(props) => {
                      const { x, y, value, index } = props;
                      const data = summary.monthlyTrend[index];
                      if (!data) return null;
                      return (
                        <g>
                          <text x={x} y={y + 18} fill="#c2410c" fontSize={12} fontWeight={800} textAnchor="middle">
                            {`${(value / 10000).toFixed(0)}만`}
                          </text>
                          <text x={x} y={y + 32} fill="#ea580c" fontSize={11} fontWeight={700} textAnchor="middle">
                            {`(${data.repurchaseShare.toFixed(1)}%)`}
                          </text>
                        </g>
                      );
                    }}
                  />
                </Area>



              </AreaChart>

            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center gap-2">
              <div className="w-2 h-6 rounded-full" style={{ background: 'linear-gradient(to bottom, #22c55e, #16a34a)' }}></div>
              <h3 className="text-lg font-bold">각인 서비스 성과 분석</h3>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2 text-sm">
                <span className="inline-block w-3 h-3 rounded-full" style={{ background: '#3b82f6' }}></span>
                <span className="text-on-surface-variant opacity-70">전체 매출</span>
              </div>
              <div className="flex items-center gap-2 text-sm">
                <span className="inline-block w-3 h-3 rounded-full" style={{ background: '#22c55e' }}></span>
                <span className="text-on-surface-variant opacity-70">각인 매출</span>
              </div>
              <div className="bg-green-500/10 text-green-600 px-4 py-1.5 rounded-full text-sm font-bold flex items-center gap-2">
                <span className="opacity-60 text-xs">채택률</span>
                {summary.engravingRate.toFixed(1)}%
              </div>
            </div>
          </div>
          <div style={{ height: '430px' }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={summary.monthlyTrend} margin={{ top: 30, right: 20, left: 30, bottom: 50 }}>
                <defs>
                  <linearGradient id="colorRevEngraving" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.15} />
                    <stop offset="95%" stopColor="#3b82f6" stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="colorEngraving" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#22c55e" stopOpacity={0.3} />
                    <stop offset="95%" stopColor="#22c55e" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                <XAxis
                  dataKey="month"
                  tick={{ fontSize: 11, fill: 'var(--color-on-surface-variant)', opacity: 0.6 }}
                  tickMargin={38}
                  axisLine={false}
                  tickLine={false}
                />
                <YAxis
                  tickFormatter={v => `${(v / 10000).toFixed(0)}만`}
                  tick={{ fontSize: 11, fill: 'var(--color-on-surface-variant)', opacity: 0.6 }}
                  axisLine={false}
                  tickLine={false}
                  tickMargin={15}
                  width={60}
                />
                <Tooltip
                  contentStyle={{ background: 'var(--color-surface-container)', border: '1px solid var(--color-outline-variant)', borderRadius: '12px' }}
                  formatter={(value, name) => [
                    `₩${value.toLocaleString()}`,
                    name === 'revenue' ? '전체 매출' : '각인 매출'
                  ]}
                />
                {/* 전체 매출 (파란색) - 레이블 위쪽 */}
                <Area
                  type="monotone"
                  dataKey="revenue"
                  stroke="#3b82f6"
                  strokeWidth={2.5}
                  fill="url(#colorRevEngraving)"
                  dot={{ r: 3, fill: '#3b82f6', strokeWidth: 0 }}
                  activeDot={{ r: 5 }}
                >
                  <LabelList
                    dataKey="revenue"
                    position="top"
                    fontSize={12}
                    fontWeight={700}
                    formatter={(v) => `${(v / 10000).toFixed(0)}만`}
                    style={{ fill: '#1d4ed8' }}
                  />
                </Area>
                {/* 각인 매출 (녹색) - 레이블 아래쪽 */}
                <Area
                  type="monotone"
                  dataKey="engravingRevenue"
                  stroke="#22c55e"
                  strokeWidth={2.5}
                  fill="url(#colorEngraving)"
                  dot={{ r: 3, fill: '#22c55e', strokeWidth: 0 }}
                  activeDot={{ r: 5 }}
                >
                  <LabelList
                    dataKey="engravingRevenue"
                    content={(props) => {
                      const { x, y, value, index } = props;
                      const data = summary.monthlyTrend[index];
                      if (!data || !value) return null;
                      // 월 매출 대비 각인 비중 계산
                      const share = data.revenue > 0 ? (value / data.revenue) * 100 : 0;
                      return (
                        <g>
                          <text x={x} y={y + 16} fill="#16a34a" fontSize={12} fontWeight={800} textAnchor="middle">
                            {`${(value / 10000).toFixed(0)}만`}
                          </text>
                          <text x={x} y={y + 29} fill="#22c55e" fontSize={11} fontWeight={700} textAnchor="middle">
                            {`(${share.toFixed(1)}%)`}
                          </text>
                        </g>
                      );
                    }}
                  />
                </Area>
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>


        {/* 결제 환경 분석 및 AI 인사이트 (위로 이동됨) */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          <div className="card">
            <h3 className="text-lg font-bold mb-6">결제 환경 분석 (MOBILE vs PC)</h3>
            <div style={{ height: '300px' }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={summary.deviceStats}
                    cx="50%"
                    cy="50%"
                    innerRadius={70}
                    outerRadius={90}
                    paddingAngle={5}
                    dataKey="value"
                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                  >
                    {summary.deviceStats.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip formatter={(value) => `₩${value.toLocaleString()}`} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="card bg-primary text-white flex flex-col justify-center items-center text-center p-8">
            <Sparkles size={48} className="mb-4 opacity-80" />
            <h3 className="text-xl font-bold mb-3">AI 마케팅 한줄 평</h3>
            {(() => {
              const trend = summary.monthlyTrend || [];
              if (trend.length < 2) return <p className="text-lg opacity-90">데이터가 부족합니다.</p>;

              // ── 1행: 피크·저점 ──
              const peakMonth = _.maxBy(trend, 'revenue');
              const lowMonth = _.minBy(trend, 'revenue');
              const peakAmt = `${(peakMonth.revenue / 10000).toLocaleString(undefined, { maximumFractionDigits: 0 })}만`;
              const lowAmt = `${(lowMonth.revenue / 10000).toLocaleString(undefined, { maximumFractionDigits: 0 })}만`;
              const line1 = `📈 피크 ${peakMonth.month} (${peakAmt})  /  📉 저점 ${lowMonth.month} (${lowAmt})`;

              // ── 2행: 전체 기간 트렌드 분석 ──
              const n = trend.length;
              const half = Math.floor(n / 2);
              const firstHalfAvg = _.meanBy(trend.slice(0, half), 'revenue');
              const secondHalfAvg = _.meanBy(trend.slice(half), 'revenue');
              const changeRate = ((secondHalfAvg - firstHalfAvg) / firstHalfAvg * 100).toFixed(1);
              const firstRev = trend[0].revenue;
              const lastRev = trend[n - 1].revenue;

              // 연속 상승/하락 구간 파악
              let risingStreak = 0, fallingStreak = 0;
              for (let i = n - 1; i > 0; i--) {
                if (trend[i].revenue > trend[i - 1].revenue) { if (fallingStreak > 0) break; risingStreak++; }
                else { if (risingStreak > 0) break; fallingStreak++; }
              }

              let trendMsg = '';
              if (parseFloat(changeRate) >= 10) {
                trendMsg = `후반부 평균이 전반 대비 ${changeRate}% 상승 — 성장세가 뚜렷합니다. 지금이 공격적 마케팅 적기입니다! 🚀`;
              } else if (parseFloat(changeRate) <= -10) {
                trendMsg = `후반부 평균이 전반 대비 ${Math.abs(changeRate)}% 하락 — 전반 대비 둔화 추세입니다. 리텐션 강화가 시급합니다. ⚠️`;
              } else if (risingStreak >= 2) {
                trendMsg = `최근 ${risingStreak}개월 연속 상승 중 — 전체 기간은 보합이지만 최근 모멘텀이 살아나고 있습니다. 📊`;
              } else if (fallingStreak >= 2) {
                trendMsg = `최근 ${fallingStreak}개월 연속 하락 중 — 전체 보합이지만 최근 흐름이 꺾였습니다. 원인 점검이 필요합니다. 📉`;
              } else if (lastRev > firstRev) {
                trendMsg = `기간 첫달 대비 마지막 달 매출이 높아 전반적인 우상향 흐름을 보입니다. 꾸준한 성장세를 유지하세요. 📈`;
              } else {
                trendMsg = `기간 내 등락이 반복되는 혼조 패턴입니다. 시즌별 프로모션 전략으로 저점 방어에 집중하세요. 🎯`;
              }

              return (
                <div className="space-y-3">
                  <p className="text-base font-semibold opacity-95">{line1}</p>
                  <p className="text-sm opacity-85 leading-relaxed">{trendMsg}</p>
                </div>
              );
            })()}
          </div>

        </div>

        {/* 주요 제품별 성과 (아래로 이동됨) */}
        <div className="card">
          <h3 className="text-lg font-bold mb-6 flex items-center justify-between">
            <div className="flex items-center" style={{ gap: '16px' }}>
              <span>주요 제품별 성과</span>
              <span className="text-sm font-semibold text-primary bg-primary/10 px-3 py-1 rounded-full border border-primary/20">
                총매출액: {summary.totalRevenue.toLocaleString()}원
              </span>
            </div>
            <div className="flex items-center" style={{ gap: '20px' }}>
              {startDate && endDate && (
                <div className="text-xs font-medium text-on-surface-variant flex items-center gap-2">
                  <span className="opacity-60">판매기간:</span>
                  <span className="bg-surface-container px-2 py-1 rounded-md">
                    {startDate.replace(/-/g, '.')} ~ {endDate.replace(/-/g, '.')}
                  </span>
                </div>
              )}
              <button 
                onClick={handleExportExcel}
                className="export-button"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px',
                  padding: '8px 16px',
                  backgroundColor: '#ffffff',
                  border: '1px solid #e2e8f0',
                  borderRadius: '12px',
                  boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
                  cursor: 'pointer',
                  transition: 'all 0.2s ease',
                  outline: 'none'
                }}
                onMouseOver={(e) => {
                  e.currentTarget.style.transform = 'translateY(-2px)';
                  e.currentTarget.style.boxShadow = '0 10px 15px -3px rgba(0,0,0,0.1)';
                  e.currentTarget.style.borderColor = '#10b981';
                }}
                onMouseOut={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
                  e.currentTarget.style.borderColor = '#e2e8f0';
                }}
                title="엑셀 보고서 추출"
              >
                <div style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  width: '32px',
                  height: '32px',
                  backgroundColor: '#ecfdf5',
                  borderRadius: '8px',
                  transition: 'background-color 0.2s'
                }}>
                  <FileSpreadsheet size={18} style={{ color: '#059669' }} />
                </div>
                <span style={{ 
                  fontSize: '14px', 
                  fontWeight: '700', 
                  color: '#475569',
                  letterSpacing: '-0.025em'
                }}>
                  보고서 추출
                </span>
              </button>
            </div>
          </h3>
          <div style={{ height: Math.max(440, summary.productStats.length * 45) + 'px' }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart
                layout="vertical"
                data={summary.productStats}
                margin={{ top: 20, right: 120, left: 40, bottom: 20 }}
              >


                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                <XAxis type="number" hide />
                <YAxis
                  dataKey="name"
                  type="category"
                  axisLine={false}
                  tickLine={false}
                  fontSize={13}
                  width={180}
                  tick={(props) => {
                    const { x, y, payload } = props;
                    return (
                      <g 
                        onClick={() => {
                          const product = summary.productStats.find(p => p.name === payload.value);
                          if (product) onDrillDown({ name: product.name, category: product.category });
                        }}
                        style={{ cursor: 'pointer' }}
                      >
                        <text 
                          x={x} 
                          y={y} 
                          dy={4} 
                          textAnchor="end" 
                          fill="#424656" 
                          fontSize={13} 
                          fontWeight={600}
                          className="hover:fill-primary transition-colors"
                        >
                          {payload.value}
                        </text>
                      </g>
                    );
                  }}
                />

                <Tooltip
                  formatter={(value, name, props) => {
                    if (name === 'revenue') {
                      const { cumulativeShare } = props.payload;
                      return [
                        <div key="tip">
                          <div>매출액: {Number(value || 0).toLocaleString()}원</div>
                          <div style={{ color: 'var(--primary)', fontWeight: 'bold', marginTop: '4px' }}>
                            누적 비중: {Number(cumulativeShare || 0).toFixed(1)}%
                          </div>
                        </div>,
                        '상세 정보'
                      ];
                    }
                    return [Number(value || 0).toLocaleString(), name];
                  }}
                />

                <Bar 
                  dataKey="revenue" 
                  radius={[0, 4, 4, 0]} 
                  barSize={24}
                  style={{ cursor: 'pointer' }}
                  onClick={(data) => {
                    if (data) {
                      onDrillDown({ name: data.name, category: data.category });
                    }
                  }}
                >


                  {summary.productStats.map((entry, index) => {
                    const pc = entry.prevCumulative;
                    let fillColor = '#ef4444'; // 90%~ (Red)
                    if (pc < 70) fillColor = 'var(--primary)'; // 0~70% (Blue)
                    else if (pc < 80) fillColor = '#f97316'; // 70~80% (Orange)
                    else if (pc < 90) fillColor = '#22c55e'; // 80~90% (Light Green)
                    else fillColor = '#ef4444'; // 90%~ (Red)
                    return <Cell key={`cell-${index}`} fill={fillColor} />;
                  })}
                  <LabelList
                    dataKey="revenue"
                    position="right"
                    content={(props) => {
                      const { x, y, width, height, value, index } = props;
                      const entry = summary.productStats[index];
                        const pc = entry?.prevCumulative ?? 0;
                        const share = entry?.share || 0;
                        
                        let textColor = '#ef4444'; // 90%~ (Red)
                        if (pc < 70) textColor = '#424656'; // 0~70% (Dark Grey)
                        else if (pc < 80) textColor = '#c2410c'; // 70~80% (Dark Orange)
                        else if (pc < 90) textColor = '#16a34a'; // 80~90% (Dark Green)
                        
                        return (
                          <text
                            x={x + width + 10}
                            y={y + height / 2}
                            dy={4}
                            fill={textColor}
                          fontSize={13}
                          fontWeight={700}
                          style={{ cursor: 'pointer' }}
                          onClick={(e) => {
                            e.stopPropagation();
                            onDrillDown({ name: entry.name, category: entry.category });
                          }}
                        >
                          {`${Number(Math.round(value / 10000) || 0).toLocaleString()}만 (${Number(share || 0).toFixed(1)}%)`}
                        </text>

                      );
                    }}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

      </div>
    </div>
  );
}

function ProductView({ summary, startDate, endDate }) {
  if (!summary || !summary.categoryStats) return null;

  // 매출 순으로 정렬하되 '기타' 카테고리는 항상 최하단에 배치
  const sortedCategories = React.useMemo(() => {
    return [...(summary.categoryStats || [])]
      .sort((a, b) => {
        if (a.name === '기타') return 1;
        if (b.name === '기타') return -1;
        return b.revenue - a.revenue;
      })
      .map(c => c.name);
  }, [summary.categoryStats]);

  const [selectedCategory, setSelectedCategory] = React.useState(sortedCategories[0] || '');

  React.useEffect(() => {
    if (sortedCategories.length > 0 && (!selectedCategory || !sortedCategories.includes(selectedCategory))) {
      setSelectedCategory(sortedCategories[0]);
    }
  }, [sortedCategories]);

  const products = summary.categoryProductMap[selectedCategory] || [];

  return (
    <div className="animate-fade-in">
      <h1 className="text-3xl font-bold mb-2">카테고리별 매출 분석</h1>
      <p className="text-on-surface-variant mb-10">상품군 카테고리별 수익성 및 시장 점유율 분석 데이터입니다.</p>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 mb-12">
        <div className="lg:col-span-2 card">
          <h3 className="text-lg font-bold mb-6 flex items-center justify-between">
            <span>전체 카테고리 매출 비교</span>
            <span className="text-xs font-normal text-on-surface-variant bg-surface-container px-2 py-1 rounded-md">
              {summary.dateRange.start} ~ {summary.dateRange.end}
            </span>
          </h3>
          <div style={{ height: Math.max(460, (summary.categoryStats || []).length * 50) + 'px' }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart layout="vertical" data={summary.categoryStats || []} margin={{ right: 90 }}>
                <XAxis type="number" hide />
                <YAxis dataKey="name" type="category" width={120} fontSize={12} axisLine={false} tickLine={false} fontWeight={600} />
                <Tooltip formatter={(value) => [`₩${Number(value || 0).toLocaleString()}`, '매출액']} />
                <Bar dataKey="revenue" fill="var(--primary)" radius={[0, 4, 4, 0]}>
                  <LabelList dataKey="revenue" position="right" fontSize={11} formatter={(v) => `₩${(Number(v || 0) / 10000).toFixed(0)}만`} style={{ fill: '#424656', fontWeight: 600 }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card">
          <h3 className="text-lg font-bold mb-6">카테고리별 매출 비중</h3>
          <div style={{ height: '460px' }}>
            <ResponsiveContainer width="100%" height="100%">
              {(() => {
                // 매출 비중이 너무 작은 카테고리(4% 미만)는 가독성을 위해 '기타'로 통합
                const total = (summary.categoryStats || []).reduce((sum, item) => sum + (item.revenue || 0), 0);
                if (total === 0) return (
                  <div className="flex items-center justify-center h-full text-on-surface-variant opacity-60">
                    분석할 데이터가 없습니다.
                  </div>
                );
                const mainCategories = [];
                let othersRevenue = 0;

                (summary.categoryStats || []).forEach(item => {
                  if (item.revenue / total >= 0.04 && item.name !== '기타') {
                    mainCategories.push(item);
                  } else {
                    othersRevenue += (item.revenue || 0);
                  }
                });

                if (othersRevenue > 0) {
                  mainCategories.push({ name: '기타', revenue: othersRevenue });
                }

                // 매출 높은 순 정렬 (단, 기타는 항상 마지막)
                const sortedCategories = mainCategories.filter(c => c.name !== '기타').sort((a, b) => b.revenue - a.revenue);
                const otherCategory = mainCategories.find(c => c.name === '기타');
                if (otherCategory) sortedCategories.push(otherCategory);

                // 커스텀 범례용 데이터 직접 생성 (Recharts 기본 정렬 완벽 차단)
                const customPayload = sortedCategories.map((entry, index) => ({
                  value: entry.name,
                  color: COLORS[index % COLORS.length],
                  payload: entry
                }));

                // 한 줄에 하나씩 보여주는 세로형 범례 렌더러 (하단 중앙 정렬)
                const renderCustomLegend = () => {
                  return (
                    <div style={{ display: 'flex', justifyContent: 'center', marginTop: '24px' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        {customPayload.map((entry, index) => {
                          const percent = total > 0 ? (Number(entry.payload.revenue || 0) / total * 100).toFixed(1) : '0.0';
                          return (
                            <div key={`item-${index}`} style={{ display: 'flex', alignItems: 'center', fontSize: '13px', color: 'var(--on-surface-variant)' }}>
                              <div style={{ width: '12px', height: '12px', backgroundColor: entry.color, marginRight: '8px', borderRadius: '3px', flexShrink: 0 }} />
                              <span style={{ fontWeight: 600, whiteSpace: 'nowrap', color: 'var(--on-surface)', minWidth: '70px', textAlign: 'left' }}>{entry.value}</span>
                              <span style={{ marginLeft: '4px', opacity: 0.8, whiteSpace: 'nowrap' }}>({percent}%)</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                };

                return (
                  <PieChart>
                    <Pie
                      data={sortedCategories}
                      dataKey="revenue"
                      cx="50%"
                      cy="45%"
                      innerRadius={55}
                      outerRadius={90}
                      label={({ name, percent }) => {
                        const p = (percent || 0) * 100;
                        return `${name} ${isNaN(p) ? '0.0' : p.toFixed(1)}%`;
                      }}
                      labelLine={{ stroke: 'var(--outline-variant)', strokeWidth: 1 }}
                    >
                      {sortedCategories.map((e, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                    </Pie>
                    <Tooltip formatter={(value) => [`₩${Number(value || 0).toLocaleString()}`, '매출액']} />
                    <Legend verticalAlign="bottom" align="center" content={renderCustomLegend} />
                  </PieChart>
                );
              })()}
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* 카테고리 제품 순위 섹션 */}
      <h3 className="text-2xl font-bold mt-12 mb-2">카테고리 상세 매출 순위</h3>
      <p className="text-on-surface-variant mb-12">선택한 카테고리 내 제품별 매출 순위와 판매 성과를 상세히 확인할 수 있습니다.</p>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: '24px' }}>
        {/* 좌측: 카테고리 버튼 목록 */}
        <div className="card" style={{ height: 'fit-content', padding: '20px' }}>
          <p className="text-sm text-on-surface-variant mb-4 font-bold">카테고리 선택</p>
          <div className="flex flex-col gap-2">
            {sortedCategories.map(cat => (
              <button
                key={cat}
                onClick={() => setSelectedCategory(cat)}
                style={{
                  textAlign: 'left',
                  padding: '12px 16px',
                  borderRadius: '10px',
                  fontSize: '14px',
                  fontWeight: selectedCategory === cat ? 700 : 500,
                  transition: 'all 0.2s',
                  backgroundColor: selectedCategory === cat ? 'var(--primary)' : 'transparent',
                  color: selectedCategory === cat ? 'white' : 'var(--on-surface)',
                  border: selectedCategory === cat ? 'none' : '1px solid var(--outline-variant)'
                }}
                className={selectedCategory === cat ? '' : 'hover:bg-surface-container-high'}
              >
                {cat}
              </button>
            ))}
          </div>
        </div>

        {/* 우측: 제품 랭킹 카드 */}
        <div className="card">
          <div className="flex justify-between items-center mb-8">
            <h3 className="text-xl font-bold flex items-center gap-6">
              <span>
                <span className="text-primary mr-3">[{selectedCategory}]</span>
                제품별 매출 순위
              </span>
              <span className="text-xs font-normal text-on-surface-variant bg-surface-container px-2 py-1 rounded-md">
                {summary.dateRange.start} ~ {summary.dateRange.end}
              </span>
            </h3>
            <span className="text-sm text-on-surface-variant bg-surface-container px-3 py-1 rounded-full">
              총 {products.length}개 품목
            </span>
          </div>

          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th style={{ width: '60px' }}>순위</th>
                  <th>제품명</th>
                  <th className="text-right">판매수량</th>
                  <th className="text-right">누적 매출액</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, idx) => (
                  <tr key={p.name} className="hover:bg-surface-container-low transition-colors">
                    <td>
                      <span className={`
                        flex items-center justify-center w-8 h-8 rounded-full text-xs font-bold
                        ${idx < 3 ? 'bg-primary text-white' : 'bg-surface-container text-on-surface-variant'}
                      `}>
                        {idx + 1}
                      </span>
                    </td>
                    <td className="font-medium text-primary">{p.name}</td>
                    <td className="text-right">{Number(p.qty || 0).toLocaleString()}개</td>
                    <td className="text-right">
                      <span className="font-bold text-lg">₩{Number(p.revenue || 0).toLocaleString()}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {products.length === 0 && (
            <div className="py-20 text-center text-on-surface-variant opacity-60">
              데이터가 없습니다.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function SkuView({ summary, drillDownInfo, onClearDrillDown, startDate, endDate, onBack }) {
  if (!summary) return null;

  // 카테고리 목록 (매출 순, 기타는 맨 뒤)
  const categories = React.useMemo(() => {
    return [...(summary.categoryStats || [])]
      .sort((a, b) => {
        if (a.name === '기타') return 1;
        if (b.name === '기타') return -1;
        return b.revenue - a.revenue;
      })
      .map(c => c.name);
  }, [summary.categoryStats]);

  const [selectedCategory, setSelectedCategory] = React.useState(categories[0] || '');
  const [selectedProduct, setSelectedProduct] = React.useState(null);

  // 선택된 카테고리의 총 매출액 계산
  const categoryTotalRevenue = React.useMemo(() => {
    const products = summary.categoryProductMap[selectedCategory] || [];
    return _.sumBy(products, 'revenue');
  }, [summary.categoryProductMap, selectedCategory]);

  // 드릴다운 정보가 있을 경우 초기값 설정
  React.useEffect(() => {
    if (drillDownInfo) {
      setSelectedCategory(drillDownInfo.category);
      setSelectedProduct(drillDownInfo.name);
    }
  }, [drillDownInfo]);

  React.useEffect(() => {
    if (categories.length > 0 && (!selectedCategory || !categories.includes(selectedCategory))) {
      setSelectedCategory(categories[0]);
    }
  }, [categories]);

  // 카테고리 변경 시 제품 선택 초기화
  const handleCategoryChange = (cat) => {
    setSelectedCategory(cat);
    setSelectedProduct(null);
  };

  const products = summary.categoryProductMap[selectedCategory] || [];
  const optionData = selectedProduct ? (summary.productOptionMap[selectedProduct] || { sizeStats: [], colorStats: [] }) : null;

  const handleExportCategoryExcel = () => {
    if (!products || products.length === 0) return;
    
    const workbook = XLSX.utils.book_new();

    // 1. 카테고리별 제품 목록 시트 구성
    const title = [[`[${selectedCategory}] 카테고리 제품 성과 보고서`]];
    const period = [[`분석 기간: ${(startDate || summary.dateRange.start).replace(/-/g, '.')} ~ ${(endDate || summary.dateRange.end).replace(/-/g, '.')}`]];
    const categoryRev = [[`카테고리 총 매출액: ${categoryTotalRevenue.toLocaleString()}원`]];
    const emptyRow = [[]];
    const headers = [['순위', '제품명', '판매수량', '매출액(원)', '매출비중(%)']];
    
    const dataRows = products.map((p, idx) => [
      idx + 1,
      p.name,
      p.qty,
      p.revenue,
      `${categoryTotalRevenue > 0 ? ((p.revenue / categoryTotalRevenue) * 100).toFixed(1) : 0}%`
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([...title, ...period, ...categoryRev, ...emptyRow, ...headers, ...dataRows]);

    // 스타일 적용 (카테고리 시트)
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[addr]) continue;
        let s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { vertical: 'center', horizontal: 'center' } };
        if (R === 0) s = { font: { name: '맑은 고딕', sz: 16, bold: true, underline: true }, alignment: { horizontal: 'left' } };
        else if (R <= 2) s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'left' } };
        else if (R === 4) {
          s = { font: { name: '맑은 고딕', sz: 10, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "333333" } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };
        } else if (R >= 5) {
          s.border = { top: { style: 'thin', color: { rgb: "CCCCCC" } }, bottom: { style: 'thin', color: { rgb: "CCCCCC" } }, left: { style: 'thin', color: { rgb: "CCCCCC" } }, right: { style: 'thin', color: { rgb: "CCCCCC" } } };
          if (C === 2 || C === 3) { s.alignment = { horizontal: 'right' }; worksheet[addr].z = '#,##0'; }
        }
        worksheet[addr].s = s;
      }
    }
    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
    worksheet['!cols'] = [{ wch: 8 }, { wch: 50 }, { wch: 12 }, { wch: 18 }, { wch: 12 }];

    XLSX.utils.book_append_sheet(workbook, worksheet, '카테고리성과');

    // 2. 제품이 선택된 상태라면 옵션 상세 시트들 추가
    if (selectedProduct && optionData) {
      const applyOptionStyles = (ws, headerRow) => {
        const r = XLSX.utils.decode_range(ws['!ref']);
        for (let row = r.s.r; row <= r.e.r; ++row) {
          for (let col = r.s.c; col <= r.e.c; ++col) {
            const addr = XLSX.utils.encode_cell({ r: row, c: col });
            if (!ws[addr]) continue;
            let s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { vertical: 'center', horizontal: 'center' } };
            
            if (row === 0) s = { font: { name: '맑은 고딕', sz: 14, bold: true }, alignment: { horizontal: 'left' } };
            else if (row === 1 || row === 2) s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'left' } };
            else if (row === headerRow) {
              s = { font: { name: '맑은 고딕', sz: 10, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "333333" } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };
            } else if (row > headerRow) {
              s.border = { top: { style: 'thin', color: { rgb: "CCCCCC" } }, bottom: { style: 'thin', color: { rgb: "CCCCCC" } }, left: { style: 'thin', color: { rgb: "CCCCCC" } }, right: { style: 'thin', color: { rgb: "CCCCCC" } } };
              if (col === 1 || col === 2) { s.alignment = { horizontal: 'right' }; ws[addr].z = '#,##0'; }
            }
            ws[addr].s = s;
          }
        }
        ws['!cols'] = [{ wch: 25 }, { wch: 12 }, { wch: 18 }, { wch: 12 }];
      };

      const periodRow = [`분석 기간: ${(startDate || summary.dateRange.start).replace(/-/g, '.')} ~ ${(endDate || summary.dateRange.end).replace(/-/g, '.')}`];

      if (optionData.sizeStats.length > 0) {
        const prodTotalRev = _.sumBy(optionData.sizeStats, 'revenue');
        const sizeTitle = [[`[${selectedProduct}] 사이즈별 판매 현황`]];
        const sizeRev = [[`제품 총 매출액: ${prodTotalRev.toLocaleString()}원`]];
        const sizeData = [['사이즈 옵션', '판매수량', '매출액(원)', '비중(%)'], ...optionData.sizeStats.map(s => [s.name, s.qty, s.revenue, `${((s.revenue / prodTotalRev) * 100).toFixed(1)}%`])];
        
        const wsSize = XLSX.utils.aoa_to_sheet([...sizeTitle, [periodRow[0]], ...sizeRev, [[]], ...sizeData]);
        applyOptionStyles(wsSize, 4);
        XLSX.utils.book_append_sheet(workbook, wsSize, '사이즈별');
      }

      if (optionData.colorStats.length > 0) {
        const prodTotalRev = _.sumBy(optionData.colorStats, 'revenue');
        const colorTitle = [[`[${selectedProduct}] 컬러별 판매 현황`]];
        const colorRev = [[`제품 총 매출액: ${prodTotalRev.toLocaleString()}원`]];
        const colorData = [['컬러 옵션', '판매수량', '매출액(원)', '비중(%)'], ...optionData.colorStats.map(c => [c.name, c.qty, c.revenue, `${((c.revenue / prodTotalRev) * 100).toFixed(1)}%`])];
        
        const wsColor = XLSX.utils.aoa_to_sheet([...colorTitle, [periodRow[0]], ...colorRev, [[]], ...colorData]);
        applyOptionStyles(wsColor, 4);
        XLSX.utils.book_append_sheet(workbook, wsColor, '컬러별');
      }
    }

    const fileName = selectedProduct 
      ? `통합보고서_${selectedCategory}_${selectedProduct.replace(/\//g, '_')}.xlsx`
      : `카테고리보고서_${selectedCategory}.xlsx`;
    
    try {
      const { ipcRenderer } = window.require('electron');
      const excelBase64 = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
      ipcRenderer.invoke('save-excel', { content: excelBase64, fileName }).then(res => {
        if (res && res.success) alert('성공적으로 저장되었습니다!');
      });
    } catch (e) { XLSX.writeFile(workbook, fileName); }
  };

  // 2. 제품 옵션 상세 엑셀 내보내기
  const handleExportOptionExcel = () => {
    if (!selectedProduct || !optionData) return;
    
    const workbook = XLSX.utils.book_new();
    const applyStyles = (ws, headerRow) => {
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[addr]) continue;
          let s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { vertical: 'center', horizontal: 'center' } };
          
          if (R === 0) s = { font: { name: '맑은 고딕', sz: 14, bold: true }, alignment: { horizontal: 'left' } };
          else if (R === 1 || R === 2) s = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'left' } };
          else if (R === headerRow) {
            s = { font: { name: '맑은 고딕', sz: 10, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "333333" } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };
          } else if (R > headerRow) {
            s.border = { top: { style: 'thin', color: { rgb: "CCCCCC" } }, bottom: { style: 'thin', color: { rgb: "CCCCCC" } }, left: { style: 'thin', color: { rgb: "CCCCCC" } }, right: { style: 'thin', color: { rgb: "CCCCCC" } } };
            if (C === 1 || C === 2) { s.alignment = { horizontal: 'right' }; ws[addr].z = '#,##0'; }
          }
          ws[addr].s = s;
        }
      }
      ws['!cols'] = [{ wch: 25 }, { wch: 12 }, { wch: 18 }, { wch: 12 }];
    };

    const periodStr = `분석 기간: ${(startDate || summary.dateRange.start).replace(/-/g, '.')} ~ ${(endDate || summary.dateRange.end).replace(/-/g, '.')}`;

    if (optionData.sizeStats.length > 0) {
      const prodTotalRev = _.sumBy(optionData.sizeStats, 'revenue');
      const sizeTitle = [[`[${selectedProduct}] 사이즈별 판매 현황`]];
      const sizeRev = [[`제품 총 매출액: ${prodTotalRev.toLocaleString()}원`]];
      const sizeData = [['사이즈 옵션', '판매수량', '매출액(원)', '비중(%)'], ...optionData.sizeStats.map(s => [s.name, s.qty, s.revenue, `${((s.revenue / prodTotalRev) * 100).toFixed(1)}%`])];
      
      const wsSize = XLSX.utils.aoa_to_sheet([...sizeTitle, [periodStr], ...sizeRev, [[]], ...sizeData]);
      applyStyles(wsSize, 4);
      XLSX.utils.book_append_sheet(workbook, wsSize, '사이즈별');
    }

    if (optionData.colorStats.length > 0) {
      const prodTotalRev = _.sumBy(optionData.colorStats, 'revenue');
      const colorTitle = [[`[${selectedProduct}] 컬러별 판매 현황`]];
      const colorRev = [[`제품 총 매출액: ${prodTotalRev.toLocaleString()}원`]];
      const colorData = [['컬러 옵션', '판매수량', '매출액(원)', '비중(%)'], ...optionData.colorStats.map(c => [c.name, c.qty, c.revenue, `${((c.revenue / prodTotalRev) * 100).toFixed(1)}%`])];
      
      const wsColor = XLSX.utils.aoa_to_sheet([...colorTitle, [periodStr], ...colorRev, [[]], ...colorData]);
      applyStyles(wsColor, 4);
      XLSX.utils.book_append_sheet(workbook, wsColor, '컬러별');
    }

    const fileName = `옵션현황_${(selectedProduct || '미선택').replace(/\//g, '_')}.xlsx`;
    try {
      const { ipcRenderer } = window.require('electron');
      const excelBase64 = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
      ipcRenderer.invoke('save-excel', { content: excelBase64, fileName }).then(res => {
        if (res && res.success) alert('성공적으로 저장되었습니다!');
      });
    } catch (e) { XLSX.writeFile(workbook, fileName); }
  };

  return (
    <div className="animate-fade-in">
      <div className="print-hidden">
        <h1 className="text-3xl font-bold mb-2">제품 옵션별 상세 현황</h1>
        <p className="text-on-surface-variant mb-10">카테고리 및 제품을 선택하면 사이즈·색상별 판매 현황을 확인할 수 있습니다.</p>
      </div>



      {/* 카테고리 + 제품 목록 */}
      <div className="print-hidden" style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: '24px' }}>
        {/* 좌측: 카테고리 버튼 */}
        <div className="card" style={{ height: 'fit-content', padding: '20px' }}>
          <p className="text-sm text-on-surface-variant mb-4 font-bold">카테고리 선택</p>
          <div className="flex flex-col gap-2">
            {categories.map(cat => (
              <button
                key={cat}
                onClick={() => handleCategoryChange(cat)}
                style={{
                  textAlign: 'left', padding: '12px 16px', borderRadius: '10px', fontSize: '14px',
                  fontWeight: selectedCategory === cat ? 700 : 500, transition: 'all 0.2s',
                  backgroundColor: selectedCategory === cat ? THEME_COLOR : 'transparent',
                  color: selectedCategory === cat ? 'white' : 'var(--on-surface)',
                  border: selectedCategory === cat ? 'none' : '1px solid var(--outline-variant)', cursor: 'pointer'
                }}
              >
                {cat}
              </button>
            ))}
          </div>
        </div>

        {/* 우측: 제품 목록 */}
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '24px' }}>
            <div>
              <h3 className="text-xl font-bold" style={{ marginBottom: '8px' }}>
                <span style={{ color: THEME_COLOR }} className="mr-2">[{selectedCategory}]</span> 제품 목록
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <span className="text-sm text-on-surface-variant bg-surface-container px-3 py-1 rounded-full">
                  총 {products.length}개 품목
                </span>
                <span style={{ fontSize: '15px', fontWeight: '700', color: '#000000', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Calendar size={16} /> {startDate || summary.dateRange.start} ~ {endDate || summary.dateRange.end}
                </span>
              </div>
            </div>
            <div style={{ textAlign: 'right', display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '12px' }}>
              <div>
                <p style={{ fontSize: '12px', color: '#666', marginBottom: '4px' }}>카테고리 총 매출액</p>
                <p style={{ fontSize: '20px', fontWeight: '800', color: THEME_COLOR }}>
                  ₩{Number(categoryTotalRevenue || 0).toLocaleString()}
                </p>
              </div>
              <button
                onClick={handleExportCategoryExcel}
                style={{
                  display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 16px',
                  backgroundColor: '#fff', color: '#374151', borderRadius: '10px',
                  fontSize: '13px', fontWeight: '700', border: '1px solid #e5e7eb',
                  cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
                  transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)', outline: 'none'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'translateY(-2px)';
                  e.currentTarget.style.boxShadow = '0 6px 12px rgba(0,0,0,0.08)';
                  e.currentTarget.style.borderColor = '#10b981';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
                  e.currentTarget.style.borderColor = '#e5e7eb';
                }}
              >
                <div style={{
                  width: '24px', height: '24px', backgroundColor: '#ecfdf5',
                  borderRadius: '6px', display: 'flex', alignItems: 'center',
                  justifyContent: 'center', color: '#10b981'
                }}>
                  <FileSpreadsheet size={14} />
                </div>
                리스트 추출
              </button>
            </div>
          </div>
          <p className="text-sm text-on-surface-variant mb-4 opacity-70">제품을 클릭하면 아래에서 사이즈·색상별 판매 현황을 확인할 수 있습니다.</p>
          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th style={{ width: '50px' }}>순위</th>
                  <th>제품명</th>
                  <th className="text-right">매출비중</th>
                  <th className="text-right">판매수량</th>
                  <th className="text-right">매출액</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, idx) => (
                  <tr
                    key={p.name}
                    onClick={() => setSelectedProduct(selectedProduct === p.name ? null : p.name)}
                    style={{ cursor: 'pointer', backgroundColor: selectedProduct === p.name ? THEME_BG : '' }}
                    className="hover:bg-surface-container-low transition-colors"
                  >
                    <td>
                      <span className={`flex items-center justify-center w-8 h-8 rounded-full text-xs font-bold ${idx < 3 ? '' : 'bg-surface-container text-on-surface-variant'}`} style={{ backgroundColor: idx < 3 ? THEME_COLOR : undefined, color: idx < 3 ? 'white' : undefined }}>
                        {idx + 1}
                      </span>
                    </td>
                    <td className="font-medium" style={{ color: THEME_COLOR, fontWeight: selectedProduct === p.name ? 800 : 500 }}>{p.name}</td>
                    <td className="text-right font-medium" style={{ color: '#666' }}>
                      {categoryTotalRevenue > 0 ? ((Number(p.revenue || 0) / categoryTotalRevenue) * 100).toFixed(1) : 0}%
                    </td>
                    <td className="text-right">{Number(p.qty || 0).toLocaleString()}개</td>
                    <td className="text-right"><span className="font-bold">₩{Number(p.revenue || 0).toLocaleString()}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
            {products.length === 0 && (
              <div className="py-16 text-center text-on-surface-variant opacity-60">데이터가 없습니다.</div>
            )}
          </div>
        </div>
      </div>

      {/* 제품 선택 시: 하단 사이즈/색상 카드 */}
      {selectedProduct && optionData && (
        <div className="mt-8">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-xl font-bold flex items-center">
              <span className="text-primary mr-2">[{selectedCategory}]</span>
              <span className="text-on-surface">{selectedProduct}</span>
              <span className="text-on-surface-variant font-normal mx-3">—</span>
              <span className="text-on-surface-variant font-normal">옵션별 판매 현황</span>
            </h3>

            <div style={{ display: 'flex', gap: '10px' }}>
              <button
                onClick={handleExportOptionExcel}
                style={{
                  display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 16px',
                  backgroundColor: '#fff', color: '#374151', borderRadius: '10px',
                  fontSize: '13px', fontWeight: '700', border: '1px solid #e5e7eb',
                  cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
                  transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)', outline: 'none'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'translateY(-2px)';
                  e.currentTarget.style.boxShadow = '0 6px 12px rgba(0,0,0,0.08)';
                  e.currentTarget.style.borderColor = '#10b981';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
                  e.currentTarget.style.borderColor = '#e5e7eb';
                }}
                className="print-hidden"
              >
                <div style={{
                  width: '24px', height: '24px', backgroundColor: '#ecfdf5',
                  borderRadius: '6px', display: 'flex', alignItems: 'center',
                  justifyContent: 'center', color: '#10b981'
                }}>
                  <FileSpreadsheet size={14} />
                </div>
                Excel 추출
              </button>


            </div>
          </div>

          {!optionData ? (
            <div className="card flex items-center justify-center py-40 opacity-50">
              <div className="text-center">
                <Package size={48} className="mx-auto mb-4" />
                <p>좌측 목록에서 분석할 제품을 선택해주세요.</p>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 print:grid-cols-1 print:gap-12">
              {/* 사이즈별 판매 현황 */}
              <div className="card flex flex-col">
                <h3 className="text-lg font-bold mb-6 flex items-center justify-between">
                  <span>사이즈별 판매 현황</span>
                  <span className="text-xs font-normal text-on-surface-variant">
                    {summary.dateRange.start} ~ {summary.dateRange.end}
                  </span>
                </h3>

                {!optionData.sizeStats || optionData.sizeStats.length === 0 ? (
                  <p className="text-on-surface-variant text-sm py-8 text-center opacity-60">사이즈 옵션 데이터가 없습니다.</p>
                ) : (
                  <div className="flex flex-col flex-1">
                    <div style={{ height: '220px' }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={optionData.sizeStats} margin={{ right: 60 }}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} fontSize={12} />
                          <YAxis axisLine={false} tickLine={false} fontSize={12} />
                          <Tooltip formatter={(v) => [`${Number(v || 0).toLocaleString()}개`, '판매수량']} />
                          <Bar dataKey="qty" fill="var(--primary)" radius={[4, 4, 0, 0]}>
                            <LabelList dataKey="qty" position="top" fontSize={11} formatter={(v) => `${Number(v || 0)}개`} style={{ fill: 'var(--on-surface-variant)', fontWeight: 600 }} />
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="table-container flex-1 flex flex-col justify-between" style={{ marginTop: '12px' }}>
                      <table>
                        <thead><tr><th>사이즈</th><th className="text-right">판매비중</th><th className="text-right">판매수량</th><th className="text-right">매출액</th></tr></thead>
                        <tbody>
                          {optionData.sizeStats.map((s, i) => {
                            const totalSizeQty = _.sumBy(optionData.sizeStats, 'qty');
                            const percent = totalSizeQty > 0 ? ((s.qty / totalSizeQty) * 100).toFixed(1) : 0;
                            return (
                              <tr key={i}>
                                <td className="font-medium">{s.name}</td>
                                <td className="text-right font-medium" style={{ color: THEME_COLOR }}>{percent}%</td>
                                <td className="text-right">{Number(s.qty || 0).toLocaleString()}개</td>
                                <td className="text-right font-bold">₩{Number(s.revenue || 0).toLocaleString()}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="border-t-2 border-outline-variant font-bold bg-surface-container-low">
                            <td colSpan="2" className="py-3">총합계</td>
                            <td className="text-right py-3">{Number(_.sumBy(optionData.sizeStats, 'qty') || 0).toLocaleString()}개</td>
                            <td className="text-right py-3" style={{ color: THEME_COLOR }}>₩{Number(_.sumBy(optionData.sizeStats, 'revenue') || 0).toLocaleString()}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
              </div>

              {/* 색상별 판매 현황 */}
              <div className="card flex flex-col">
                <h3 className="text-lg font-bold mb-6 flex items-center justify-between">
                  <span>색상별 판매 현황</span>
                  <span className="text-xs font-normal text-on-surface-variant">
                    {summary.dateRange.start} ~ {summary.dateRange.end}
                  </span>
                </h3>

                {!optionData.colorStats || optionData.colorStats.length === 0 ? (
                  <p className="text-on-surface-variant text-sm py-8 text-center opacity-60">색상 옵션 데이터가 없습니다.</p>
                ) : (
                  <div className="flex flex-col flex-1">
                    <div style={{ height: '220px' }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={optionData.colorStats.slice(0, 5)}
                            dataKey="qty"
                            cx="50%"
                            cy="50%"
                            innerRadius={55}
                            outerRadius={85}
                            label={({ cx, cy, midAngle, innerRadius, outerRadius, value, index, x, y, name, percent }) => {
                              const p = (percent || 0) * 100;
                              return (
                                <text x={x} y={y} fill="var(--on-surface-variant)" textAnchor={x > cx ? 'start' : 'end'} dominantBaseline="central" fontSize={11} fontWeight={500}>
                                  {`${name} ${isNaN(p) ? '0.0' : p.toFixed(1)}%`}
                                </text>
                              );
                            }}
                            labelLine={{ stroke: 'var(--outline-variant)', strokeWidth: 1 }}
                          >
                            {optionData.colorStats.slice(0, 5).map((e, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                          </Pie>
                          <Tooltip formatter={(v) => [`${Number(v || 0).toLocaleString()}개`, '판매수량']} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="table-container flex-1 flex flex-col justify-between" style={{ marginTop: '12px' }}>
                      <table>
                        <thead><tr><th>색상</th><th className="text-right">판매비중</th><th className="text-right">판매수량</th><th className="text-right">매출액</th></tr></thead>
                        <tbody>
                          {optionData.colorStats.map((c, i) => {
                            const totalColorQty = _.sumBy(optionData.colorStats, 'qty');
                            const percent = totalColorQty > 0 ? ((c.qty / totalColorQty) * 100).toFixed(1) : 0;
                            return (
                              <tr key={i}>
                                <td>
                                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
                                    <span style={{ width: '10px', height: '10px', borderRadius: '3px', backgroundColor: COLORS[i % COLORS.length], flexShrink: 0, display: 'inline-block' }} />
                                    {c.name}
                                  </span>
                                </td>
                                <td className="text-right font-medium" style={{ color: THEME_COLOR }}>{percent}%</td>
                                <td className="text-right">{Number(c.qty || 0).toLocaleString()}개</td>
                                <td className="text-right font-bold">₩{Number(c.revenue || 0).toLocaleString()}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="border-t-2 border-outline-variant font-bold bg-surface-container-low">
                            <td colSpan="2" className="py-3">총합계</td>
                            <td className="text-right py-3">{Number(_.sumBy(optionData.colorStats, 'qty') || 0).toLocaleString()}개</td>
                            <td className="text-right py-3" style={{ color: THEME_COLOR }}>₩{Number(_.sumBy(optionData.colorStats, 'revenue') || 0).toLocaleString()}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 월별 매출 추이 (1단 카드) */}
          {(() => {
            if (!optionData) return null;

            // 전체 월매출 Map 생성 (month → totalRevenue)
            const totalByMonth = {};
            (summary.monthlyTrend || []).forEach(m => { totalByMonth[m.month] = m.revenue; });

            // 제품 월매출에 전체 월매출 및 비중(%) 합산
            const mergedTrend = (optionData.monthlyTrend || []).map(m => ({
              ...m,
              totalRevenue: totalByMonth[m.month] || 0,
              share: totalByMonth[m.month] > 0
                ? ((m.revenue / totalByMonth[m.month]) * 100)
                : 0
            }));

            return (
              <div className="card mt-8">
                <h3 className="text-lg font-bold mb-2 flex items-center justify-between">
                  <span>월별 매출 추이</span>
                  <span className="text-xs font-normal text-on-surface-variant">
                    {summary.dateRange.start} ~ {summary.dateRange.end}
                  </span>
                </h3>
                {/* 범례 */}
                <div className="flex items-center gap-6 mb-4 text-xs text-on-surface-variant">
                  <span className="flex items-center gap-1.5">
                    <span style={{ display:'inline-block', width:20, height:3, backgroundColor: THEME_COLOR, borderRadius:2 }}/>
                    {selectedProduct}
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span style={{ display:'inline-block', width:20, height:3, backgroundColor:'#94a3b8', borderRadius:2 }}/>
                    전체 월매출
                  </span>
                </div>
                <div style={{ height: '340px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={mergedTrend} margin={{ top: 30, right: 40, left: 0, bottom: 0 }}>
                      <defs>
                        <linearGradient id="gradProduct" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={THEME_COLOR} stopOpacity={0.12}/>
                          <stop offset="95%" stopColor={THEME_COLOR} stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="gradTotal" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#94a3b8" stopOpacity={0.10}/>
                          <stop offset="95%" stopColor="#94a3b8" stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="month" fontSize={11} tickLine={false} axisLine={false} dy={10} />

                      {/* 왼쪽 Y축: 전체 월매출 스케일 */}
                      <YAxis
                        yAxisId="left"
                        orientation="left"
                        fontSize={10}
                        tickLine={false}
                        axisLine={false}
                        tickFormatter={(v) => `${(Number(v || 0) / 10000).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}만`}
                        tick={{ fill: '#94a3b8' }}
                        width={55}
                      />
                      {/* 오른쪽 Y축: 선택 제품 매출 스케일 */}
                      <YAxis
                        yAxisId="right"
                        orientation="right"
                        fontSize={10}
                        tickLine={false}
                        axisLine={false}
                        tickFormatter={(v) => `${(Number(v || 0) / 10000).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}만`}
                        tick={{ fill: THEME_COLOR }}
                        width={55}
                      />

                      <Tooltip
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 12px rgba(0,0,0,0.1)', fontSize: 12 }}
                        formatter={(v, name) => {
                          if (name === 'revenue') return [`${Number(v || 0).toLocaleString()}원`, selectedProduct];
                          if (name === 'totalRevenue') return [`${Number(v || 0).toLocaleString()}원`, '전체 월매출'];
                          return [v, name];
                        }}
                        labelFormatter={(label, payload) => {
                          const d = payload?.[0]?.payload;
                          const shareStr = d && d.share != null ? ` (비중 ${Number(d.share).toFixed(1)}%)` : '';
                          return `${label}${shareStr}`;
                        }}
                      />
                      {/* 전체 월매출 (배경 Area, 왼쪽 축 기준) */}
                      <Area
                        yAxisId="left"
                        type="monotone"
                        dataKey="totalRevenue"
                        stroke="#94a3b8"
                        strokeWidth={2}
                        strokeDasharray="5 3"
                        fillOpacity={1}
                        fill="url(#gradTotal)"
                        dot={false}
                      >
                        <LabelList
                          dataKey="totalRevenue"
                          position="top"
                          offset={10}
                          fontSize={10}
                          fontWeight={600}
                          fill="#94a3b8"
                          formatter={(v) => `${(Number(v || 0) / 10000).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}만`}
                        />
                      </Area>

                      {/* 제품 월매출 (전경 Area, 오른쪽 축 기준) */}
                      <Area
                        yAxisId="right"
                        type="monotone"
                        dataKey="revenue"
                        stroke={THEME_COLOR}
                        strokeWidth={3}
                        fillOpacity={1}
                        fill="url(#gradProduct)"
                        dot={{ r: 4, fill: THEME_COLOR, stroke: '#fff', strokeWidth: 2 }}
                      >
                        {/* 매출액 레이블 */}
                        <LabelList
                          dataKey="revenue"
                          position="top"
                          offset={12}
                          fontSize={10}
                          fontWeight={600}
                          fill={THEME_COLOR}
                          formatter={(v) => `${(Number(v || 0) / 10000).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}만`}
                        />
                      </Area>
                      {/* 비중(%) 레이블 전용 투명 라인 (오른쪽 축 기준) */}
                      <Line
                        yAxisId="right"
                        type="monotone"
                        dataKey="share"
                        stroke="transparent"
                        dot={false}
                        legendType="none"
                        tooltipType="none"
                        isAnimationActive={false}
                      >
                        <LabelList
                          dataKey="share"
                          position="top"
                          offset={32}
                          fontSize={10}
                          fontWeight={700}
                          fill="#e11d48"
                          formatter={(v) => `${Number(v || 0).toFixed(1)}%`}
                        />
                      </Line>

                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            );
          })()}
        </div>

      )}

    </div>
  );
}

function GuideView({ summary, onUpload, loading }) {
  const guideData = [
    {
      title: "대시보드 (Dashboard)",
      icon: <LayoutDashboard size={28} />,
      desc: "매출 흐름과 KPI를 한눈에 파악하세요.",
      steps: [
        "분석 기간 설정: 원하는 기간을 선택하여 특정 시점의 성과를 필터링할 수 있습니다.",
        "매출 트렌드: 월별 매출 변화와 재구매 매출 비중을 차트로 확인합니다.",
        "핵심 지표: 총 매출, 주문 건수, 재구매율, 각인 채택률을 실시간으로 집계합니다."
      ]
    },
    {
      title: "카테고리 분석 (Category)",
      icon: <BarChart3 size={28} />,
      desc: "인기 제품군과 효율적인 포트폴리오를 분석합니다.",
      steps: [
        "카테고리 점유율: 전체 매출 중 각 카테고리가 차지하는 비중을 확인합니다.",
        "성과 비교: 판매 수량 대비 매출액이 높은 '효자 카테고리'를 식별합니다.",
        "상세 리스트: 각 카테고리 클릭 시 분류에 속한 제품별 성과를 보여줍니다."
      ]
    },
    {
      title: "옵션별 현황 (SKU Status)",
      icon: <ClipboardList size={28} />,
      desc: "색상, 사이즈별 판매 데이터를 상세히 분석합니다.",
      steps: [
        "베스트 옵션 파악: 가장 많이 팔리는 조합을 확인하여 재고 관리에 활용합니다.",
        "드릴다운 분석: 대시보드 제품 클릭 시 해당 제품의 옵션 현황으로 이동합니다.",
        "옵션 트렌드: 월별 상승 옵션을 파악하여 시즌별 기획에 참고하세요."
      ]
    }
  ];

  return (
    <div style={{ animation: 'fadeIn 0.5s ease-out', paddingBottom: '48px', maxWidth: '1200px', margin: '0 auto' }}>
      <div style={{ 
        display: 'flex', 
        justifyContent: 'space-between', 
        alignItems: 'center', 
        marginBottom: '64px',
        padding: '32px',
        backgroundColor: '#ffffff',
        borderRadius: '24px',
        boxShadow: '0 4px 20px rgba(0,0,0,0.05)',
        border: '1px solid #f0f0f0'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '24px' }}>
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center', 
            width: '64px', 
            height: '64px', 
            borderRadius: '50%', 
            backgroundColor: '#f0f4ff', 
            color: '#3b82f6'
          }}>
            <Sparkles size={32} />
          </div>
          <div>
            <h1 style={{ fontSize: '32px', fontWeight: '800', color: '#1a1a1a', margin: 0 }}>판매 분석 마스터 가이드</h1>
            <p style={{ fontSize: '16px', color: '#666', marginTop: '4px', margin: 0 }}>데이터 대시보드를 200% 활용하기 위한 핵심 매뉴얼입니다.</p>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
          {loading ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', color: '#3b82f6', fontWeight: '700' }}>
              <div className="animate-spin" style={{ width: '24px', height: '24px', border: '3px solid #eff6ff', borderTopColor: '#3b82f6', borderRadius: '50%' }}></div>
              데이터 분석 중...
            </div>
          ) : (
            <label style={{ 
              display: 'flex', 
              alignItems: 'center', 
              gap: '12px', 
              backgroundColor: '#3b82f6', 
              color: '#ffffff', 
              padding: '18px 36px', 
              borderRadius: '20px', 
              fontSize: '20px', 
              fontWeight: '800', 
              cursor: 'pointer',
              boxShadow: '0 10px 25px rgba(59, 130, 246, 0.4)',
              transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)'
            }}
            className="hover-scale-lift"
            >
              <Upload size={26} />
              <span>{summary ? '새 데이터 분석' : '데이터 업로드하여 분석 시작'}</span>
              <input type="file" accept=".csv" onChange={onUpload} style={{ display: 'none' }} />
            </label>
          )}
        </div>
      </div>

      {/* 스마트스토어 데이터 준비 가이드 (최상단으로 이동) */}
      <div style={{ 
        marginBottom: '48px', 
        backgroundColor: '#ffffff', 
        borderRadius: '24px', 
        padding: '40px', 
        boxShadow: '0 10px 30px rgba(0,0,0,0.1)',
        border: '2px solid #3b82f6'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '32px' }}>
          <div style={{ padding: '12px', borderRadius: '12px', backgroundColor: '#eff6ff', color: '#3b82f6' }}>
            <FileSpreadsheet size={28} />
          </div>
          <div>
            <h3 style={{ fontSize: '24px', fontWeight: '800', color: '#1a1a1a', margin: 0 }}>네이버 스마트스토어 데이터 준비 가이드</h3>
            <p style={{ color: '#3b82f6', fontWeight: '600', marginTop: '4px' }}>정확한 분석을 위해 반드시 아래 순서대로 파일을 준비해주세요.</p>
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '24px' }}>
          {[
            { title: "STEP 01. 엑셀 다운로드", content: "스마트스토어 접속 > 판매관리 > 구매확정 내역 > 기간 선택 후 [검색] > [엑셀다운로드] (비번 설정 권장)" },
            { title: "STEP 02. 파일 편집 허용", content: "다운로드된 엑셀 파일을 실행한 후, 상단의 [편집 사용] 버튼을 클릭합니다." },
            { title: "STEP 03. CSV 형식으로 저장", content: "[다른 이름으로 저장] 클릭 > 파일 형식을 'CSV UTF-8 (쉼표로 분리)'로 반드시 선택하여 저장합니다." },
            { title: "STEP 04. 데이터 업로드", content: "프로그램 실행 후 저장한 CSV 파일을 업로드하면 즉시 판매 분석이 시작됩니다." }
          ].map((step, idx) => (
            <div key={idx} style={{ 
              padding: '24px', 
              borderRadius: '20px', 
              backgroundColor: '#f8f9fa', 
              border: '1px solid #e9ecef',
              position: 'relative'
            }}>
              <div style={{ 
                fontSize: '12px', 
                fontWeight: '800', 
                color: '#3b82f6', 
                backgroundColor: '#ffffff', 
                display: 'inline-block', 
                padding: '4px 12px', 
                borderRadius: '8px', 
                marginBottom: '12px',
                boxShadow: '0 2px 4px rgba(0,0,0,0.05)'
              }}>
                {step.title}
              </div>
              <p style={{ fontSize: '14px', color: '#333', lineHeight: '1.6', margin: 0 }}>
                {step.content}
              </p>
            </div>
          ))}
        </div>

        <div style={{ 
          marginTop: '24px', 
          padding: '16px 20px', 
          borderRadius: '12px', 
          backgroundColor: '#fffbeb', 
          border: '1px solid #fef3c7', 
          display: 'flex', 
          gap: '12px', 
          alignItems: 'center' 
        }}>
          <span style={{ fontSize: '20px' }}>💡</span>
          <p style={{ fontSize: '13px', color: '#92400e', margin: 0 }}>
            <strong>주의사항:</strong> 일반 Excel(.xlsx) 파일은 인식이 되지 않을 수 있습니다. 반드시 <strong>CSV UTF-8</strong> 형식을 확인해주세요!
          </p>
        </div>
      </div>

      {/* 기능별 상세 가이드 (하단으로 이동) */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '32px' }}>
        {guideData.map((guide, idx) => (
          <div key={idx} style={{ 
            backgroundColor: '#ffffff', 
            borderRadius: '24px', 
            padding: '32px', 
            boxShadow: '0 4px 20px rgba(0,0,0,0.08)', 
            display: 'flex', 
            flexDirection: 'column' 
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '32px' }}>
              <div style={{ padding: '12px', borderRadius: '12px', backgroundColor: '#f8f9fa', color: '#3b82f6' }}>
                {guide.icon}
              </div>
              <h3 style={{ fontSize: '20px', fontWeight: '700', color: '#1a1a1a', margin: 0 }}>{guide.title}</h3>
            </div>
            
            <p style={{ fontSize: '16px', fontWeight: '600', color: '#3b82f6', marginBottom: '32px', lineHeight: '1.5' }}>
              {guide.desc}
            </p>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '24px', marginTop: 'auto' }}>
              {guide.steps.map((step, sIdx) => {
                const [label, content] = step.split(':');
                return (
                  <div key={sIdx} style={{ display: 'flex', gap: '16px', alignItems: 'flex-start' }}>
                    <span style={{ 
                      flexShrink: 0, 
                      width: '24px', 
                      height: '24px', 
                      borderRadius: '50%', 
                      backgroundColor: '#eff6ff', 
                      color: '#3b82f6', 
                      display: 'flex', 
                      alignItems: 'center', 
                      justifyContent: 'center', 
                      fontSize: '12px', 
                      fontWeight: 'bold',
                      marginTop: '2px'
                    }}>
                      {sIdx + 1}
                    </span>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                      <strong style={{ fontSize: '14px', color: '#1a1a1a' }}>{label}</strong>
                      <span style={{ fontSize: '13px', color: '#666', lineHeight: '1.6' }}>{content}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>

      <div style={{ 
        marginTop: '80px', 
        padding: '48px', 
        borderRadius: '32px', 
        backgroundColor: '#f8f9fa', 
        border: '1px solid #e9ecef', 
        textAlign: 'center' 
      }}>
        <h4 style={{ fontSize: '24px', fontWeight: '700', color: '#1a1a1a', marginBottom: '16px' }}>도움이 필요하신가요?</h4>
        <p style={{ fontSize: '18px', color: '#666', marginBottom: '32px' }}>데이터 분석 중 궁금한 점이 생기면 언제든 다니엘에게 물어보세요!</p>
        <div style={{ display: 'flex', justifyContent: 'center', gap: '16px' }}>
          <button style={{ 
            padding: '12px 32px', 
            borderRadius: '9999px', 
            backgroundColor: '#3b82f6', 
            color: '#ffffff', 
            fontWeight: '700', 
            border: 'none', 
            cursor: 'pointer',
            boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)'
          }}>다니엘에게 질문하기</button>
          <button style={{ 
            padding: '12px 32px', 
            borderRadius: '9999px', 
            backgroundColor: '#ffffff', 
            border: '1px solid #dee2e6', 
            color: '#495057', 
            fontWeight: '700', 
            cursor: 'pointer'
          }}>문의하기</button>
        </div>
      </div>
    </div>
  );
}






function SummaryCard({ title, value, trend, up, neutral, titleClassName }) {
  return (
    <div className="card stat-card">
      <p className={titleClassName || "stat-label"}>{title}</p>
      <p className="stat-value" style={{ fontSize: '33px' }}>{value}</p>
      {!neutral && (
        <div className={`stat-trend ${up ? 'trend-up' : 'trend-down'}`} style={{ fontSize: '18px' }}>
          {up ? <TrendingUp size={18} /> : <TrendingDown size={18} />}
          {trend}
        </div>
      )}
      {neutral && (
        <div className="stat-trend" style={{ color: 'var(--on-surface-variant)', fontSize: '18px' }}>
          {trend}
        </div>
      )}
    </div>
  );
}

function InsightAction({ title, desc }) {
  return (
    <div className="card group hover:border-primary transition-all cursor-pointer">
      <div className="flex justify-between items-start">
        <div>
          <h4 className="font-bold mb-3 group-hover:text-primary transition-colors">{title}</h4>
          <p className="text-sm text-on-surface-variant leading-relaxed">{desc}</p>
        </div>
        <ChevronRight size={20} className="text-outline group-hover:translate-x-1 transition-transform" />
      </div>
    </div>
  );
}

export default App;
