require 'spec_helper'

office_versions = YAML.load_file(Dir.pwd + '/spec/office_versions.yml')

describe 'msoffice::servicepack', type: :define do
  let :hiera_data do
    {
      windows_deployment_root: '\\test-server\packages',
      company_name: 'Example Inc',
      office_username: 'Joe'
    }
  end

  describe 'installing with unknown version' do
    let :title do
      'servicepack for unknown version'
    end
    let :params do
      { version: 'xxx', sp: '1' }
    end

    it do
      expect { is_expected.to contain_exec('install-sp') }.
        to raise_error(Puppet::Error, %r{The version agrument specified does not match a valid version of office})
    end
  end

  describe 'installing with unknown sp' do
    let :title do
      'servicepack for unknown sp version'
    end
    let :params do
      { version: '2010', sp: '5' }
    end

    it do
      expect { is_expected.to contain_exec('install-sp') }.
        to raise_error(Puppet::Error, %r{The service pack specified not supported})
    end
  end

  describe 'incorrect arch' do
    let :title do
      'servicepack with incorrect arch'
    end
    let :params do
      { arch: 'fubar', version: '2010', sp: '1' }
    end

    it do
      expect { is_expected.to contain_exec('install-sp') }.
        to raise_error(Puppet::Error, %r{The arch argument specified does not match x86 or x64})
    end
  end

  %w[1 2 3].each do |sp|
    version = '2003'
    office_num = office_versions[version]['version']
    build = office_versions[version]['service_packs'][sp]['build']
    setup = office_versions[version]['service_packs'][sp]['setup']
    describe "installing office #{version} SP#{sp}" do
      let :title do
        "SP#{sp} for office #{version}"
      end
      let(:params) do
        {
          version: version,
          sp: sp,
          deployment_root: '\\test-server\\packages'
        }
      end

      it do
        is_expected.to contain_exec('install-sp').with(
          'command' => "& \"\\test-server\\packages\\OFFICE#{office_num}\\SPs\\#{setup}\" /q /norestart",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\#{office_num}.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
      end
    end
  end

  %w[1 2 3].each do |sp|
    version = '2007'
    office_num = office_versions[version]['version']
    build = office_versions[version]['service_packs'][sp]['build']
    setup = office_versions[version]['service_packs'][sp]['setup']
    describe "installing office #{version} SP#{sp}" do
      let :title do
        "SP#{sp} for office #{version}"
      end
      let(:params) do
        {
          version: version,
          sp: sp,
          deployment_root: '\\test-server\packages'
        }
      end

      it do
        is_expected.to contain_exec('install-sp').with(
          'command' => "& \"\\test-server\\packages\\OFFICE#{office_num}\\SPs\\#{setup}\" /q /norestart",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\#{office_num}.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
      end
    end
  end

  %w[1 2].each do |sp|
    version = '2010'
    office_num = office_versions[version]['version']
    build = office_versions[version]['service_packs'][sp]['build']
    setup = office_versions[version]['service_packs'][sp]['setup']['x86']
    describe "installing office #{version} SP#{sp}" do
      let :title do
        "SP#{sp} for office #{version}"
      end
      let(:params) do
        {
          arch: 'x86',
          version: version,
          sp: sp,
          deployment_root: '\\test-server\packages'
        }
      end

      it do
        is_expected.to contain_exec('install-sp').with(
          'command' => "& \"\\test-server\\packages\\OFFICE#{office_num}\\SPs\\x86\\#{setup}\" /q /norestart",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\#{office_num}.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
      end
    end
  end
end
