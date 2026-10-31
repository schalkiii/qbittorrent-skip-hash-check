#!/usr/bin/env python3
"""
qBittorrent WebAPI 鉴权模块（同时被 skip-hash-check 与 tracker-tagger 复用）

支持两种鉴权方式，优先使用 API token：
1. API token（qBittorrent >= 5.2.0 / WebAPI >= v2.14.1）
   - 请求头：Authorization: Bearer <token>
   - token 格式：32 字符，"qbt_" 前缀 + 28 位随机字符
   - 特性：无状态、不依赖 Cookie、无需登录往返；与用户名密码互斥且优先级更高
   - 限制：不能用于 auth/ 端点（login/logout 会返回 403），也不能访问非 /api/v2/ 路径
   - 注意：实测 X-API-Key 头在 v5.3.0 上返回 403，故统一使用 Bearer
2. 用户名 + 密码（传统 Cookie 会话）
   - 先 POST /api/v2/auth/login 获取会话 Cookie
   - 兼容新旧版本：旧版成功返回 200 + "Ok."，新版（4.6+）返回 204 且响应体为空
"""

import sys

import requests

API_KEY_PREFIX = 'qbt_'
API_KEY_LENGTH = 32
REQUEST_TIMEOUT = (10, 60)  # (连接超时, 读取超时)


class AuthError(RuntimeError):
    """鉴权失败"""


def validate_token(token):
    """校验 token 格式，返回是否形如 qbt_ + 28 位字符（共 32 字符）"""
    if not token:
        return False
    return token.startswith(API_KEY_PREFIX) and len(token) == API_KEY_LENGTH


class QBClient:
    """
    统一的 qBittorrent WebAPI 客户端。

    用法：
        client = QBClient.build(host, token, username, password)
        client.get('/api/v2/torrents/info').json()
    """

    def __init__(self, base_url, session, auth_mode):
        self.base_url = base_url.rstrip('/')
        self.session = session
        self.auth_mode = auth_mode  # 'token' 或 'password'

    # ---------- 构造 ----------

    @classmethod
    def build(cls, host, token=None, username=None, password=None):
        """
        按 token > 用户名密码 的优先级建立已鉴权的客户端。
        token 存在时完全跳过登录流程；否则回落到 Cookie 登录。
        """
        base_url = host.rstrip('/')
        session = requests.Session()
        session.headers.update({'Referer': base_url})  # 新版 qb 的 CSRF 校验要求 Referer 与 Host 一致

        if token:
            cls._check_token_auth(session, base_url, token)
            return cls(base_url, session, 'token')

        if not (username and password):
            raise AuthError('缺少凭据：请提供 --token，或同时提供 --username 与 --password')

        cls._login_with_password(session, base_url, username, password)
        return cls(base_url, session, 'password')

    # ---------- 鉴权实现 ----------

    @staticmethod
    def _check_token_auth(session, base_url, token):
        """
        验证 token 可用性。

        token 不能用于 auth/ 端点（服务端直接 403），因此改为请求一个开销极小的公开性 API。
        这里选 /api/v2/app/version：响应仅一个版本字符串，是成本最低的探测点。
        """
        if not validate_token(token):
            raise AuthError(
                f'API token 格式不正确：应为 "{API_KEY_PREFIX}" 前缀 + 28 位字符，共 {API_KEY_LENGTH} 字符，'
                f'当前为 {len(token)} 字符'
            )

        session.headers.update({'Authorization': f'Bearer {token}'})

        url = base_url + '/api/v2/app/version'
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as e:
            raise AuthError(f'无法连接 qBittorrent（{base_url}）：{e}')

        if r.status_code == 403:
            raise AuthError(
                'API token 鉴权失败（HTTP 403）。可能原因：\n'
                '  1. token 已失效或被重新生成（旧 token 会立即作废）\n'
                '  2. qBittorrent 版本低于 5.2.0，不支持 API token\n'
                '  3. WebUI 中未生成 API key\n'
                '可回退使用 --username/--password。'
            )
        if not (200 <= r.status_code < 300):
            raise AuthError(f'API token 鉴权失败：HTTP {r.status_code} {r.text[:200]}')

        return r.text.strip()

    @staticmethod
    def _login_with_password(session, base_url, username, password):
        """用户名密码登录。兼容 200 + "Ok."（旧版）与 204 空响应（4.6+）"""
        url = base_url + '/api/v2/auth/login'
        try:
            r = session.post(url, data={'username': username, 'password': password},
                             timeout=REQUEST_TIMEOUT)
        except requests.RequestException as e:
            raise AuthError(f'无法连接 qBittorrent（{base_url}）：{e}')

        if r.status_code == 403:
            # qb 在连续失败达到阈值后会封禁来源 IP
            raise AuthError('登录失败（HTTP 403）：用户名或密码错误，或该 IP 因多次失败已被临时封禁')
        if not (200 <= r.status_code < 300):
            raise AuthError(f'登录失败：HTTP {r.status_code} {r.text[:200]}')

        if not any(c.startswith('QBT_SID') for c in session.cookies.keys()):
            raise AuthError(f'登录失败：服务端未下发会话 Cookie（HTTP {r.status_code}）')
        return True

    # ---------- 请求封装 ----------

    def url(self, path):
        return self.base_url + path

    def request(self, method, path, **kwargs):
        kwargs.setdefault('timeout', REQUEST_TIMEOUT)
        r = self.session.request(method, self.url(path), **kwargs)
        if r.status_code == 403 and self.auth_mode == 'token':
            raise AuthError('请求被拒绝（HTTP 403）：API token 可能已失效，或该端点不支持 token 鉴权')
        return r

    def get(self, path, **kwargs):
        return self.request('GET', path, **kwargs)

    def post(self, path, **kwargs):
        return self.request('POST', path, **kwargs)

    def logout(self):
        """仅 Cookie 模式需要登出，用于主动释放服务端会话。token 模式为无状态，无需处理。"""
        if self.auth_mode != 'password':
            return False
        try:
            return self.post('/api/v2/auth/logout').status_code == 200
        except Exception:
            return False


def add_auth_args(parser):
    """为 argparse 解析器添加鉴权相关参数"""
    parser.add_argument('--token', '-t', default=None,
                        help='qBittorrent API token（>=5.2.0），形如 qbt_xxx；优先于用户名密码')
    parser.add_argument('--username', '-u', default=None, help='WebUI 用户名（未提供 token 时必填）')
    parser.add_argument('--password', '-p', default=None, help='WebUI 密码（未提供 token 时必填）')


def build_client_from_args(args):
    """由命令行参数构建客户端，统一鉴权失败的输出格式"""
    try:
        return QBClient.build(args.host, getattr(args, 'token', None),
                              args.username, args.password)
    except AuthError as e:
        print(f'鉴权失败: {e}')
        sys.exit(1)
